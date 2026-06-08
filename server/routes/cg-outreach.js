'use strict';

// ── CG Outreach Agent Routes ──────────────────────────────────────────────────
// Admin-only (gated client-side by password). No server-side auth needed.
// All routes mount at /api/cg via server.js

const express  = require('express');
const router   = express.Router();
const Anthropic = require('@anthropic-ai/sdk');
const dns      = require('dns').promises;

// Lazy-init so missing env vars at module-load time don't crash the server
function getAI()     { return new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY }); }

const MODEL      = 'claude-sonnet-4-20250514';
const FROM_EMAIL = process.env.SENDER_EMAIL || 'john@comptongroupllc.com';
const DAILY_LIMIT = 10; // warmup default; adjustable per-run via run-outreach body.limit

// ── Local targeting config ───────────────────────────────────────────────────
// Prospect discovery is constrained to these two metros + surrounding suburbs.
// Anything outside is rejected at insert time.
const TARGET_METROS = {
  atlanta: {
    label: 'Atlanta, GA metro',
    state: 'GA',
    // Representative metro/suburb cities used to validate a prospect's location.
    cities: ['atlanta','marietta','alpharetta','roswell','sandy springs','smyrna','dunwoody',
      'decatur','kennesaw','duluth','lawrenceville','suwanee','johns creek','peachtree corners',
      'norcross','tucker','stone mountain','douglasville','woodstock','canton','acworth','cumming',
      'snellville','stockbridge','mcdonough','newnan','peachtree city','fayetteville','conyers',
      'lithonia','college park','east point','austell','powder springs','buford','sugar hill',
      'mableton','grayson','loganville','dallas','hiram','villa rica','union city','hampton',
      'jonesboro','riverdale','forest park','morrow','ellenwood','clarkston','chamblee','brookhaven','vinings'],
  },
  birmingham: {
    label: 'Birmingham, AL metro',
    state: 'AL',
    cities: ['birmingham','hoover','vestavia hills','mountain brook','homewood','trussville',
      'bessemer','hueytown','pelham','alabaster','helena','gardendale','fultondale','irondale',
      'leeds','moody','pell city','chelsea','calera','clay','pinson','center point','mccalla',
      'springville','warrior','fairfield','tarrant','midfield','adamsville','columbiana','montevallo'],
  },
};

// Business categories the finder targets. Admin picks one or more; default = all Home Services.
const PROSPECT_CATEGORIES = {
  home_services: {
    label: 'Home Services',
    types: ['HVAC','plumbing','roofing','landscaping','electrical','garage doors','pest control'],
  },
  local_professional: {
    label: 'Local Professional Services',
    types: ['dental','chiropractic','law firms','real estate','med spas','accounting'],
  },
};

// Returns the metro key for a given city/state, or null if outside both metros.
function resolveMetro(city, state) {
  const c = (city || '').toLowerCase().trim();
  const s = (state || '').toUpperCase().trim();
  for (const [key, m] of Object.entries(TARGET_METROS)) {
    if (m.cities.includes(c)) return key;
    // Fallback: allow city match even if state string is messy, but if state is given it must match.
    if (m.cities.some(city2 => c.includes(city2)) && (!s || s === m.state)) return key;
  }
  return null;
}

// ── Microsoft Graph email transport (app-only / client credentials) ──────────
// Sends real mail through Microsoft 365 as SENDER_EMAIL. All credentials come
// from process.env (never hardcoded). Token is cached in memory and reused
// until ~60s before expiry.
let _graphToken = null;       // { value, expiresAt (ms epoch) }

async function getGraphToken() {
  const now = Date.now();
  if (_graphToken && _graphToken.expiresAt - 60_000 > now) {
    return _graphToken.value; // still valid (with 60s safety margin)
  }
  const tenant = process.env.GRAPH_TENANT_ID;
  const clientId = process.env.GRAPH_CLIENT_ID;
  const clientSecret = process.env.GRAPH_CLIENT_SECRET;
  if (!tenant || !clientId || !clientSecret) {
    throw new Error('Microsoft Graph not configured (missing GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET)');
  }

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const resp = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  const text = await resp.text();
  if (!resp.ok) {
    throw new Error(`Graph token request failed (${resp.status}): ${text}`);
  }
  const json = JSON.parse(text);
  _graphToken = {
    value: json.access_token,
    expiresAt: Date.now() + (Number(json.expires_in) || 3600) * 1000,
  };
  return _graphToken.value;
}

// sendEmail(to, subject, bodyHtml) → resolves on HTTP 202, throws on any failure
// with the response body included so it surfaces in the Outreach Log.
async function sendEmail(to, subject, bodyHtml) {
  const sender = process.env.SENDER_EMAIL || FROM_EMAIL;
  const token = await getGraphToken();

  const payload = {
    message: {
      subject,
      body: { contentType: 'HTML', content: bodyHtml },
      toRecipients: [{ emailAddress: { address: to } }],
    },
    saveToSentItems: true,
  };

  const resp = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sender)}/sendMail`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (resp.status === 202) {
    return { ok: true, status: 202 };
  }
  const errBody = await resp.text().catch(() => '');
  const msg = `Graph sendMail failed (${resp.status}) to ${to}: ${errBody}`;
  console.error('[cg/sendEmail]', msg);
  throw new Error(msg);
}

let pool; // injected by server.js
function setPool(p) { pool = p; }

// ── Email discovery + verification ────────────────────────────────────────────
// Find the single BEST real, PUBLISHED contact email for a business, then verify
// it (format + DNS MX) before it can ever be stored. We never fabricate or
// pattern-guess an address — a wrong/guessed email bounces and damages our
// sending domain's reputation.

const EMAIL_RX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function isValidEmailFormat(email) {
  return typeof email === 'string' && EMAIL_RX.test(email.trim());
}

// Resolve MX records for an email's domain. Returns true only if the domain can
// actually receive mail. Falls back to checking A records is NOT done — no MX
// means no reliable mail delivery, so we reject.
async function domainHasMx(email) {
  const at = email.lastIndexOf('@');
  if (at < 0) return false;
  const domain = email.slice(at + 1).trim().toLowerCase();
  if (!domain || domain.indexOf('.') < 0) return false;
  try {
    const mx = await dns.resolveMx(domain);
    return Array.isArray(mx) && mx.length > 0 && mx.some(r => r.exchange);
  } catch (_) {
    return false; // NXDOMAIN / no MX / lookup failure → cannot receive mail
  }
}

// Ask Claude (with web search) to find the best published email for ONE prospect.
// Returns { email, email_type, email_source } with email_type in
// 'owner'|'role'|'generic', or { email: '' } when nothing real is found.
async function discoverBestEmail(p) {
  const ownerHint = (p.contact_name || p.first_name)
    ? `The owner/primary contact appears to be "${p.contact_name || p.first_name}"${p.contact_title ? ` (${p.contact_title})` : ''}. Look for THIS person's direct email FIRST.`
    : `No specific contact name is known.`;

  const prompt = `Find the single BEST real, PUBLISHED contact email for this small local business so a founder can send them a cold outreach email that actually reaches a decision-maker.

Business: ${p.company || ''}
City/State: ${p.city || ''}, ${p.state || (p.metro === 'birmingham' ? 'AL' : p.metro === 'atlanta' ? 'GA' : '')}
Website: ${p.website || '(unknown)'}
Industry: ${p.industry || ''}
${ownerHint}

Search the business's OWN website (contact, about, team, staff pages) and reputable business listings (Google Business, Yelp, Facebook, BBB, chamber of commerce, industry directories).

Choose ONE email by this strict priority:
1. owner — the owner's or a named decision-maker's DIRECT personal business email (best chance of a reply/sale).
2. role  — the business's primary monitored mailbox or a sales@ address an owner likely reads.
3. generic — a generic info@/contact@/office@ address ONLY as a last resort.

ABSOLUTE HARD RULES:
- Only return an email that is ACTUALLY PUBLISHED on a real, verifiable source you found.
- NEVER invent an address. NEVER construct one from a pattern (do NOT build "info@"+domain or "firstname@"+domain). If that exact address is not genuinely published somewhere, do NOT return it.
- If you cannot find any real published email, return an empty string for email.

Return JSON ONLY, no markdown, exactly:
{"email":"the published email or empty string","email_type":"owner|role|generic","email_source":"short description + URL where you found it"}`;

  const response = await getAI().messages.create({
    model: MODEL,
    max_tokens: 1500,
    tools: [{ type: 'web_search_20250305', name: 'web_search' }],
    messages: [{ role: 'user', content: prompt }],
  });

  let raw = '';
  for (const block of response.content) {
    if (block.type === 'text') raw += block.text;
  }

  let out = {};
  try {
    const match = raw.match(/\{[\s\S]*\}/);
    if (match) out = JSON.parse(match[0]);
  } catch (_) { out = {}; }

  const email = (out.email || '').trim().toLowerCase();
  let type = (out.email_type || '').trim().toLowerCase();
  if (!['owner', 'role', 'generic'].includes(type)) type = 'generic';
  return { email, email_type: type, email_source: (out.email_source || '').trim() };
}

// Discover + verify + persist the best email for one prospect row.
// Returns { ok, email, email_type, email_source, reason }.
async function findAndVerifyEmailFor(p) {
  let cand;
  try {
    cand = await discoverBestEmail(p);
  } catch (e) {
    return { ok: false, reason: 'discovery_failed: ' + e.message };
  }
  if (!cand.email) return { ok: false, reason: 'no_published_email_found' };
  if (!isValidEmailFormat(cand.email)) return { ok: false, reason: 'invalid_format' };
  const mxOk = await domainHasMx(cand.email);
  if (!mxOk) return { ok: false, reason: 'no_mx_records' };

  // Verified — safe to store.
  await pool.query(
    `UPDATE cg_prospects
       SET email=$1, email_type=$2, email_verified=TRUE, email_source=$3
     WHERE id=$4`,
    [cand.email, cand.email_type, cand.email_source || '', p.id]
  );
  return { ok: true, email: cand.email, email_type: cand.email_type, email_source: cand.email_source };
}

// ── POST /api/cg/prospects/:id/find-email ─────────────────────────────────────
// Discover + verify the best published email for ONE prospect and store it.
router.post('/prospects/:id/find-email', async (req, res) => {
  try {
    const r = await pool.query(`SELECT * FROM cg_prospects WHERE id=$1`, [req.params.id]);
    if (!r.rows.length) return res.status(404).json({ error: 'Not found' });
    const result = await findAndVerifyEmailFor(r.rows[0]);
    const fresh = await pool.query(`SELECT * FROM cg_prospects WHERE id=$1`, [req.params.id]);
    res.json({ result, prospect: fresh.rows[0] });
  } catch (e) {
    console.error('[cg/find-email]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/cg/find-missing-emails ──────────────────────────────────────────
// Backfill: run discovery + verification for every prospect that does not yet
// have a verified email. Existing rows fill in without re-finding everything.
router.post('/find-missing-emails', async (req, res) => {
  try {
    const r = await pool.query(
      `SELECT * FROM cg_prospects
        WHERE email_verified IS NOT TRUE
          AND status NOT IN ('replied','converted','skipped')
        ORDER BY created_at DESC`
    );
    const targets = r.rows;
    let verified = 0;
    let failed = 0;
    const details = [];
    for (const p of targets) {
      const result = await findAndVerifyEmailFor(p);
      if (result.ok) { verified++; details.push({ company: p.company, email: result.email, type: result.email_type }); }
      else { failed++; details.push({ company: p.company, skipped: result.reason }); }
    }
    const all = await pool.query(`SELECT * FROM cg_prospects ORDER BY created_at DESC`);
    res.json({ scanned: targets.length, verified, failed, details, prospects: all.rows });
  } catch (e) {
    console.error('[cg/find-missing-emails]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/cg/prospects ─────────────────────────────────────────────────────
router.get('/prospects', async (req, res) => {
  try {
    const r = await pool.query(`SELECT * FROM cg_prospects ORDER BY created_at DESC`);
    res.json(r.rows);
  } catch (e) {
    console.error('[cg/prospects]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/cg/find-prospects ───────────────────────────────────────────────
// Local targeting: constrained to the Atlanta, GA and Birmingham, AL metros.
// Body: { categories?: string[], description?: string }
//   categories = keys from PROSPECT_CATEGORIES (default ['home_services']).
//   description = optional free-text refinement appended to the search.
router.post('/find-prospects', async (req, res) => {
  // Resolve requested categories → list of business types. Default = all Home Services.
  let categoryKeys = Array.isArray(req.body.categories) ? req.body.categories : [];
  categoryKeys = categoryKeys.filter(k => PROSPECT_CATEGORIES[k]);
  if (!categoryKeys.length) categoryKeys = ['home_services'];
  const typeList = categoryKeys.flatMap(k => PROSPECT_CATEGORIES[k].types);
  const categoryLabels = categoryKeys.map(k => PROSPECT_CATEGORIES[k].label).join(' and ');

  const refinement = (req.body.description || '').trim();
  const metroList = Object.values(TARGET_METROS).map(m => m.label).join(' and ');

  const userPrompt = `Search the web and find 15 small, local, owner-operated businesses ONLY in the ${metroList} (including surrounding suburbs).

Target these business types: ${typeList.join(', ')}.${refinement ? `\nAdditional refinement: ${refinement}` : ''}

Hard rules:
- ONLY businesses physically located in the Atlanta, GA metro or the Birmingham, AL metro (or their suburbs). Do NOT return businesses anywhere else.
- Prefer small local shops over big regional chains or franchises.

For each business capture these signals (they drive email personalization):
- company: business name
- first_name: the owner/primary contact's FIRST name only if findable, else ""
- contact_name: full contact name if findable, else ""
- contact_title: their title/role, else ""
- email: a real contact email if findable, else ""
- phone: main phone number if findable, else ""
- website: full website URL if they have one, else ""
- has_website: true if they have a real website, else false
- online_booking: true if their site offers online booking/scheduling, else false
- google_rating: their Google star rating as a number (e.g. 4.7) if findable, else null
- review_count: their Google review count as an integer if findable, else null
- city: the city they're located in
- state: "GA" or "AL"
- industry: the specific business type (e.g. "HVAC", "dental")
- ai_opportunity: ONE specific, concrete AI idea that would help THIS business based on its signals

Return a JSON array ONLY, no preamble or markdown. Each element must have exactly these keys:
company, first_name, contact_name, contact_title, email, phone, website, has_website, online_booking, google_rating, review_count, city, state, industry, ai_opportunity`;

  try {
    const response = await getAI().messages.create({
      model: MODEL,
      max_tokens: 4000,
      tools: [{
        type: 'web_search_20250305',
        name: 'web_search',
      }],
      messages: [{ role: 'user', content: userPrompt }],
    });

    // Extract text from response
    let raw = '';
    for (const block of response.content) {
      if (block.type === 'text') raw += block.text;
    }

    let prospects = [];
    try {
      const match = raw.match(/\[[\s\S]*\]/);
      if (match) prospects = JSON.parse(match[0]);
    } catch (parseErr) {
      console.error('[cg/find-prospects] parse error:', parseErr.message);
      return res.status(500).json({ error: 'AI returned unparseable JSON' });
    }

    // Insert into DB, skip duplicates by email. Reject anything outside the metros.
    let inserted = 0;
    let rejected = 0;
    const categoryLabelForRow = categoryLabels;
    for (const p of prospects) {
      if (!p.company) continue;
      const metro = resolveMetro(p.city, p.state);
      if (!metro) { rejected++; continue; } // outside Atlanta/Birmingham → reject
      try {
        const hasSite = p.has_website != null ? !!p.has_website : !!(p.website && p.website.trim());
        await pool.query(
          `INSERT INTO cg_prospects
             (company, first_name, contact_name, contact_title, email, phone, website,
              has_website, online_booking, google_rating, review_count,
              city, metro, industry, category, ai_opportunity, status)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,'pending')
           ON CONFLICT DO NOTHING`,
          [
            p.company || '',
            p.first_name || '',
            p.contact_name || '',
            p.contact_title || '',
            '', // email intentionally left blank — only a verified, published email
                // (discovered + MX-checked via find-email) is ever stored.
            p.phone || '',
            p.website || '',
            hasSite,
            p.online_booking != null ? !!p.online_booking : false,
            (p.google_rating != null && !isNaN(parseFloat(p.google_rating))) ? parseFloat(p.google_rating) : null,
            (p.review_count != null && !isNaN(parseInt(p.review_count))) ? parseInt(p.review_count) : null,
            p.city || '',
            metro,
            p.industry || '',
            categoryLabelForRow,
            p.ai_opportunity || '',
          ]
        );
        inserted++;
      } catch (rowErr) {
        console.warn('[cg/find-prospects] row insert error:', rowErr.message);
      }
    }

    const all = await pool.query(`SELECT * FROM cg_prospects ORDER BY created_at DESC`);
    res.json({ inserted, rejected, prospects: all.rows });
  } catch (e) {
    console.error('[cg/find-prospects]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── PATCH /api/cg/prospects/:id/status ───────────────────────────────────────
router.patch('/prospects/:id/status', async (req, res) => {
  const { status } = req.body;
  const allowed = ['pending', 'approved', 'sent_1', 'sent_2', 'sent_3', 'replied', 'converted', 'skipped'];
  if (!status || !allowed.includes(status)) {
    return res.status(400).json({ error: 'Invalid status' });
  }
  try {
    const r = await pool.query(
      `UPDATE cg_prospects SET status=$1 WHERE id=$2 RETURNING *`,
      [status, req.params.id]
    );
    if (!r.rows.length) return res.status(404).json({ error: 'Not found' });
    res.json(r.rows[0]);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/cg/prospects/:id/approve ───────────────────────────────────────
// Approve a prospect for send. Requires drafted emails to exist (review-before-send),
// so nothing can be queued for sending until the admin has reviewed the drafts.
router.post('/prospects/:id/approve', async (req, res) => {
  try {
    const r = await pool.query(`SELECT * FROM cg_prospects WHERE id=$1`, [req.params.id]);
    if (!r.rows.length) return res.status(404).json({ error: 'Not found' });
    const prospect = r.rows[0];

    if (!prospect.email_verified || !prospect.email) {
      return res.status(400).json({ error: 'No verified email. Run "Find email" to discover and verify a real address before approving.' });
    }

    if (!prospect.drafted || !prospect.draft_email1_body) {
      return res.status(400).json({ error: 'Draft the 3 emails and review them before approving.' });
    }

    const updated = await pool.query(
      `UPDATE cg_prospects SET status='approved' WHERE id=$1 RETURNING *`,
      [req.params.id]
    );
    res.json(updated.rows[0]);
  } catch (e) {
    console.error('[cg/approve]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── DELETE /api/cg/prospects/:id ─────────────────────────────────────────────
router.delete('/prospects/:id', async (req, res) => {
  try {
    await pool.query(`DELETE FROM cg_prospects WHERE id=$1`, [req.params.id]);
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/cg/sequences ─────────────────────────────────────────────────────
router.get('/sequences', async (req, res) => {
  try {
    const r = await pool.query(`SELECT * FROM cg_sequences ORDER BY sequence_type`);
    res.json(r.rows);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/cg/sequences ────────────────────────────────────────────────────
router.post('/sequences', async (req, res) => {
  const { sequence_type, email1_subject, email1_body, email2_subject, email2_body, email3_subject, email3_body } = req.body;
  const allowed = ['efficiency', 'replace_software', 'add_ai'];
  if (!sequence_type || !allowed.includes(sequence_type)) {
    return res.status(400).json({ error: 'Invalid sequence_type' });
  }
  try {
    const r = await pool.query(
      `INSERT INTO cg_sequences (sequence_type, email1_subject, email1_body, email2_subject, email2_body, email3_subject, email3_body)
       VALUES ($1,$2,$3,$4,$5,$6,$7)
       ON CONFLICT (sequence_type) DO UPDATE SET
         email1_subject=$2, email1_body=$3, email2_subject=$4, email2_body=$5,
         email3_subject=$6, email3_body=$7, updated_at=NOW()
       RETURNING *`,
      [sequence_type, email1_subject||'', email1_body||'', email2_subject||'', email2_body||'',
       email3_subject||'', email3_body||'']
    );
    res.json(r.rows[0]);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/cg/draft-sequence ───────────────────────────────────────────────
router.post('/draft-sequence', async (req, res) => {
  const { sequence_type } = req.body;

  const prompts = {
    efficiency: `Write a 3-email cold outreach sequence from JohnMark Compton, founder of Compton Group LLC (comptongroupllc.com), a custom AI software development company. Target: small to mid-size business owner. Angle: we build custom AI tools that automate their biggest time-wasters and operational bottlenecks, saving 10-20 hours/week. Email 1: short curious hook about one specific inefficiency in their industry. Email 2: a concrete example of an AI tool we could build for them. Email 3: soft close offering a free 20-minute discovery call. Tone: direct, founder-to-founder, no corporate fluff. Return JSON only: {"email1_subject":"...","email1_body":"...","email2_subject":"...","email2_body":"...","email3_subject":"...","email3_body":"..."}`,

    replace_software: `Write a 3-email cold outreach sequence from JohnMark Compton, founder of Compton Group LLC, targeting a small/mid-size business still using outdated software (spreadsheets, legacy systems, generic tools). Angle: we replace clunky old software with a custom AI platform built specifically for their business, often for less than their current software subscriptions. Email 1: pain point hook about outdated tools slowing them down. Email 2: what a custom-built AI platform could look like for their business. Email 3: free audit offer — we review their current stack and show what we'd build instead. Tone: conversational, no jargon, founder voice. Return JSON only: {"email1_subject":"...","email1_body":"...","email2_subject":"...","email2_body":"...","email3_subject":"...","email3_body":"..."}`,

    add_ai: `Write a 3-email cold outreach sequence from JohnMark Compton, founder of Compton Group LLC, targeting a small/mid-size business doing well but not yet using AI. Angle: we add AI to their existing workflows without rebuilding anything — think AI that handles their reports, customer follow-ups, scheduling, or data analysis automatically. Email 1: curiosity hook about one AI use case relevant to their industry. Email 2: show a specific workflow we could automate for them in 2-4 weeks. Email 3: offer a free workflow analysis — 20 minutes, we identify their top 3 AI opportunities. Tone: energetic, specific, founder-to-founder. Return JSON only: {"email1_subject":"...","email1_body":"...","email2_subject":"...","email2_body":"...","email3_subject":"...","email3_body":"..."}`,
  };

  const prompt = prompts[sequence_type];
  if (!prompt) return res.status(400).json({ error: 'Invalid sequence_type' });

  try {
    const msg = await getAI().messages.create({
      model: MODEL,
      max_tokens: 2000,
      system: `You are an expert B2B SaaS sales copywriter. Write concise, compelling cold emails. Keep each email under 150 words. Direct. No fluff. No "I hope this email finds you well." Founder-to-founder voice.`,
      messages: [{ role: 'user', content: prompt }],
    });

    let seq = {};
    try {
      const match = (msg.content[0]?.text || '').match(/\{[\s\S]*\}/);
      if (match) seq = JSON.parse(match[0]);
    } catch (parseErr) {
      return res.status(500).json({ error: 'AI returned unparseable JSON' });
    }
    res.json(seq);
  } catch (e) {
    console.error('[cg/draft-sequence]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/cg/prospects/:id/draft ─────────────────────────────────────────
// Generate a personalized 3-email sequence for ONE prospect in JohnMark's voice,
// using the prospect's captured signals. Stores into the draft_* columns and
// marks drafted=true. Does NOT send — admin reviews/edits/approves first.
const DRAFT_SYSTEM_PROMPT = `You write cold outreach emails for JohnMark, who runs a small software company (Compton Group LLC) in the Atlanta–Birmingham area building simple AI tools for local businesses.

HARD RULES:
- Short, plain-text emails that sound like a real person typed them fast.
- Use contractions. First person. Sign as "JohnMark".
- NO markdown. NO buzzwords (never "solutions", "leverage", "cutting-edge", "revolutionize"). NO corporate tone. No "I hope this email finds you well."
- One link max (and usually none).
- Email 1 must be under ~90 words.
- The offer is ALWAYS a FREE, no-pressure look — a quick 15-min call OR a local in-person visit — where JohnMark shows one specific AI idea for their business.
- NEVER pitch "custom software" and NEVER mention a price.

Use these EXACT templates, filling {{...}} per the prospect:

EMAIL 1 (opener)
Subject: quick AI idea for {{business_name}}
Hey {{first_name}},
I'm JohnMark — I run a small software company in the Atlanta–Birmingham area and I build simple AI tools for local businesses.
I was checking out {{business_name}} and noticed {{observation}}. There's a quick AI setup that could {{benefit}} — and most shops your size aren't using it yet, so it's an easy edge.
I'm not pitching some big expensive project. I'd just show you exactly what it'd look like for your business on a quick 15-min call, or swing by since I'm local. Free, no pressure.
Worth a look?
JohnMark
Compton Group LLC

EMAIL 2 (bump)
Subject: re: quick AI idea for {{business_name}}
Hey {{first_name}} — bumping this in case it slipped by.
If a call feels like a lot, I can just put together a quick 2-minute example of what the setup would do for {{business_name}} and send it over. No meeting needed — just reply "send it" and I will.
JohnMark

EMAIL 3 (close)
Subject: re: quick AI idea for {{business_name}}
Hey {{first_name}}, I'll leave it here so I'm not cluttering your inbox.
If {{pain}} ever becomes worth fixing, just reply and I'll show you what's possible — free either way. Best of luck with {{business_name}}.
JohnMark

PERSONALIZATION — pick the most relevant {{observation}}, {{benefit}}, {{pain}} from the prospect's signals/category:
- No online booking → observation: "you don't have an easy way for customers to book online"; benefit: "let people book jobs right from their phone, day or night"; pain: "missed bookings"
- Phone-only / no booking + service trade → observation: "it looks like everything runs through phone calls"; benefit: "instantly text back the calls you miss so you stop losing jobs"; pain: "missed calls"
- Strong reviews but weak/no website → observation: "you've got great reviews but not much of a web presence"; benefit: "turn that reputation into more booked jobs automatically"; pain: "leads slipping through"
You may adapt the wording slightly to fit the business, but keep the voice and structure.
{{first_name}}: use the contact's first name; if unknown, use a natural greeting ("Hey there") or the business name.

Return JSON ONLY, no markdown, with exactly these keys:
{"email1_subject":"...","email1_body":"...","email2_subject":"...","email2_body":"...","email3_subject":"...","email3_body":"..."}`;

router.post('/prospects/:id/draft', async (req, res) => {
  try {
    const r = await pool.query(`SELECT * FROM cg_prospects WHERE id=$1`, [req.params.id]);
    if (!r.rows.length) return res.status(404).json({ error: 'Not found' });
    const p = r.rows[0];

    const signals = [
      `Business name: ${p.company || ''}`,
      `Contact first name: ${p.first_name || '(unknown)'}`,
      `Category: ${p.category || ''}`,
      `Industry/type: ${p.industry || ''}`,
      `City: ${p.city || ''} (${p.metro || ''} metro)`,
      `Phone: ${p.phone || '(unknown)'}`,
      `Has website: ${p.has_website ? 'yes' : 'no'}`,
      `Online booking: ${p.online_booking ? 'yes' : 'no'}`,
      `Google rating: ${p.google_rating != null ? p.google_rating : '(unknown)'}`,
      `Review count: ${p.review_count != null ? p.review_count : '(unknown)'}`,
      `Noted AI opportunity: ${p.ai_opportunity || '(none)'}`,
    ].join('\n');

    const msg = await getAI().messages.create({
      model: MODEL,
      max_tokens: 1500,
      system: DRAFT_SYSTEM_PROMPT,
      messages: [{
        role: 'user',
        content: `Write the 3-email sequence for this prospect using their signals:\n\n${signals}`,
      }],
    });

    let seq = {};
    try {
      const match = (msg.content[0]?.text || '').match(/\{[\s\S]*\}/);
      if (match) seq = JSON.parse(match[0]);
    } catch (parseErr) {
      return res.status(500).json({ error: 'AI returned unparseable JSON' });
    }

    const updated = await pool.query(
      `UPDATE cg_prospects SET
         draft_email1_subject=$1, draft_email1_body=$2,
         draft_email2_subject=$3, draft_email2_body=$4,
         draft_email3_subject=$5, draft_email3_body=$6,
         drafted=TRUE
       WHERE id=$7 RETURNING *`,
      [seq.email1_subject || '', seq.email1_body || '',
       seq.email2_subject || '', seq.email2_body || '',
       seq.email3_subject || '', seq.email3_body || '', req.params.id]
    );
    res.json(updated.rows[0]);
  } catch (e) {
    console.error('[cg/prospects/draft]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── PUT /api/cg/prospects/:id/drafts ─────────────────────────────────────────
// Admin saves edited drafts (review-before-send). Keeps drafted=true.
router.put('/prospects/:id/drafts', async (req, res) => {
  const b = req.body || {};
  try {
    const updated = await pool.query(
      `UPDATE cg_prospects SET
         draft_email1_subject=$1, draft_email1_body=$2,
         draft_email2_subject=$3, draft_email2_body=$4,
         draft_email3_subject=$5, draft_email3_body=$6,
         drafted=TRUE
       WHERE id=$7 RETURNING *`,
      [b.email1_subject || '', b.email1_body || '',
       b.email2_subject || '', b.email2_body || '',
       b.email3_subject || '', b.email3_body || '', req.params.id]
    );
    if (!updated.rows.length) return res.status(404).json({ error: 'Not found' });
    res.json(updated.rows[0]);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/cg/outreach-log ──────────────────────────────────────────────────
router.get('/outreach-log', async (req, res) => {
  try {
    const r = await pool.query(`
      SELECT l.*, p.company, p.contact_name, p.email AS prospect_email, p.industry, p.ai_opportunity, p.status AS prospect_status
      FROM cg_outreach_log l
      JOIN cg_prospects p ON p.id = l.prospect_id
      ORDER BY l.sent_at DESC
    `);
    res.json(r.rows);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/cg/run-outreach ─────────────────────────────────────────────────
// Sends next email in sequence for all approved prospects (up to DAILY_LIMIT)
router.post('/run-outreach', async (req, res) => {
  try {
    // Daily send cap — default 10 (warmup), adjustable per run from the UI.
    let limit = parseInt(req.body && req.body.limit);
    if (!Number.isFinite(limit) || limit < 1) limit = DAILY_LIMIT;
    if (limit > 100) limit = 100;

    // Get prospects mid-sequence. Replied/converted/skipped are excluded so we
    // never email a prospect who has already replied or been marked done.
    const prospectsR = await pool.query(`
      SELECT * FROM cg_prospects
      WHERE status IN ('approved','sent_1','sent_2')
      LIMIT $1
    `, [limit]);

    const prospects = prospectsR.rows;
    if (!prospects.length) {
      return res.json({ sent: 0, message: 'No approved prospects ready for outreach' });
    }

    let sent = 0;
    const errors = [];
    const now = new Date();

    for (const p of prospects) {
      // Safety: never send to a prospect who has replied or been marked done.
      if (['replied', 'converted', 'skipped'].includes(p.status)) continue;

      // Hard gate: never send without a real, verified email. No self-fallback.
      if (!p.email_verified || !p.email || !EMAIL_RX.test(String(p.email).trim())) {
        errors.push({ company: p.company, error: 'Skipped — no verified email' });
        continue;
      }

      // Per-prospect drafted emails are the source of truth (review-before-send).
      // Determine which email to send + the timing gate.
      let emailNum = null;
      let subject  = null;
      let body     = null;
      let newStatus = null;

      if (p.status === 'approved') {
        emailNum  = 1;
        subject   = p.draft_email1_subject;
        body      = p.draft_email1_body;
        newStatus = 'sent_1';
      } else if (p.status === 'sent_1') {
        // Send email 2 only if ~3+ days since email 1
        const daysSince = p.email1_sent_at
          ? (now - new Date(p.email1_sent_at)) / 86400000
          : 99;
        if (daysSince < 3) continue;
        emailNum  = 2;
        subject   = p.draft_email2_subject;
        body      = p.draft_email2_body;
        newStatus = 'sent_2';
      } else if (p.status === 'sent_2') {
        // Send email 3 only if ~4+ days since email 2
        const daysSince = p.email2_sent_at
          ? (now - new Date(p.email2_sent_at)) / 86400000
          : 99;
        if (daysSince < 4) continue;
        emailNum  = 3;
        subject   = p.draft_email3_subject;
        body      = p.draft_email3_body;
        newStatus = 'sent_3';
      }

      if (!emailNum || !subject || !body) {
        errors.push({ company: p.company, error: `No drafted content for email ${emailNum}` });
        continue;
      }

      // Drafts are already personalized per prospect; send as-is.
      const personalSubject = subject;
      const personalBody = body;

      try {
        const html = `<div style="font-family:Arial,sans-serif;max-width:600px;line-height:1.7;color:#1a1a2e">
${personalBody.replace(/\n/g, '<br>')}
<br><br>
<hr style="border:none;border-top:1px solid #e5e7eb;margin:24px 0">
<p style="font-size:11px;color:#6b7280">
  JohnMark Compton · Founder, Compton Group LLC<br>
  <a href="https://comptongroupllc.com" style="color:#c9a84c">comptongroupllc.com</a>
</p>
</div>`;
        await sendEmail(p.email, personalSubject, html); // verified email guaranteed by the gate above

        // Log the send
        await pool.query(
          `INSERT INTO cg_outreach_log (prospect_id, sequence_type, email_number, resend_id, status)
           VALUES ($1,$2,$3,$4,'sent')`,
          [p.id, p.sequence_type || 'efficiency', emailNum, null]
        );

        // Update prospect status + timestamp
        const tsField = emailNum === 1 ? 'email1_sent_at' : emailNum === 2 ? 'email2_sent_at' : 'email3_sent_at';
        await pool.query(
          `UPDATE cg_prospects SET status=$1, ${tsField}=NOW() WHERE id=$2`,
          [newStatus, p.id]
        );

        sent++;
      } catch (sendErr) {
        console.error('[cg/run-outreach] send error for', p.company, sendErr.message);
        errors.push({ company: p.company, error: sendErr.message });
      }
    }

    res.json({ sent, total_prospects: prospects.length, errors: errors.length ? errors : undefined });
  } catch (e) {
    console.error('[cg/run-outreach]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/cg/stats ─────────────────────────────────────────────────────────
router.get('/stats', async (req, res) => {
  try {
    const [totalR, sentR, repliedR, convertedR, lastRunR, byIndustryR, recentR, bySeqR] = await Promise.all([
      pool.query(`SELECT COUNT(*) AS cnt FROM cg_prospects`),
      pool.query(`SELECT COUNT(*) AS cnt FROM cg_outreach_log`),
      pool.query(`SELECT COUNT(*) AS cnt FROM cg_prospects WHERE status='replied'`),
      pool.query(`SELECT COUNT(*) AS cnt FROM cg_prospects WHERE status='converted'`),
      pool.query(`SELECT MAX(sent_at) AS last_run FROM cg_outreach_log`),
      pool.query(`SELECT industry, COUNT(*) AS cnt FROM cg_prospects GROUP BY industry ORDER BY cnt DESC`),
      pool.query(`
        SELECT l.sent_at, l.email_number, p.company, l.sequence_type
        FROM cg_outreach_log l
        JOIN cg_prospects p ON p.id = l.prospect_id
        ORDER BY l.sent_at DESC LIMIT 10
      `),
      pool.query(`
        SELECT l.sequence_type,
          COUNT(*) FILTER (WHERE p.status='replied' OR p.status='converted') AS replies,
          COUNT(DISTINCT l.prospect_id) AS total
        FROM cg_outreach_log l
        JOIN cg_prospects p ON p.id = l.prospect_id
        GROUP BY l.sequence_type
      `),
    ]);

    const totalProspects = parseInt(totalR.rows[0].cnt) || 0;
    const totalSent      = parseInt(sentR.rows[0].cnt) || 0;
    const replied        = parseInt(repliedR.rows[0].cnt) || 0;
    const converted      = parseInt(convertedR.rows[0].cnt) || 0;
    const replyRate      = totalSent > 0 ? ((replied + converted) / totalProspects * 100).toFixed(1) : '0.0';

    // Best sequence
    let bestSeq = { name: '—', replyRate: 0 };
    for (const row of bySeqR.rows) {
      const rate = row.total > 0 ? (parseInt(row.replies) / parseInt(row.total) * 100) : 0;
      if (rate > bestSeq.replyRate) {
        bestSeq = { name: row.sequence_type, replyRate: rate.toFixed(1) };
      }
    }

    res.json({
      totalProspects,
      totalSent,
      replied,
      converted,
      replyRate,
      bestSequence: bestSeq,
      lastRun: lastRunR.rows[0]?.last_run || null,
      byIndustry: byIndustryR.rows,
      recentActivity: recentR.rows,
    });
  } catch (e) {
    console.error('[cg/stats]', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/cg/badge ─────────────────────────────────────────────────────────
router.get('/badge', async (req, res) => {
  try {
    const r = await pool.query(`SELECT COUNT(*) AS cnt FROM cg_prospects WHERE status='replied'`);
    res.json({ replied: parseInt(r.rows[0].cnt) || 0 });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/cg/test-email ───────────────────────────────────────────────────
// Admin-only test send through Microsoft Graph. Confirms Graph works before
// running real outreach. Returns the full result (or error) so the admin sees it.
router.post('/test-email', async (req, res) => {
  const { to } = req.body;
  const emailRx = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!to || !emailRx.test(to.trim())) {
    return res.status(400).json({ error: 'A valid recipient email is required' });
  }
  const recipient = to.trim();
  const subject = 'Compton Group LLC — Microsoft Graph test email';
  const html = `<div style="font-family:Arial,sans-serif;max-width:600px;line-height:1.7;color:#1a1a2e">
<p>This is a test email sent from the CG Outreach Agent via the Microsoft Graph API.</p>
<p>If you are reading this, Microsoft 365 sending is wired up correctly. ✅</p>
<hr style="border:none;border-top:1px solid #e5e7eb;margin:24px 0">
<p style="font-size:11px;color:#6b7280">
  JohnMark Compton · Founder, Compton Group LLC<br>
  <a href="https://comptongroupllc.com" style="color:#c9a84c">comptongroupllc.com</a>
</p>
</div>`;
  try {
    const result = await sendEmail(recipient, subject, html);
    console.log('[cg/test-email] sent to', recipient, result);
    res.json({ success: true, to: recipient, sender: process.env.SENDER_EMAIL || FROM_EMAIL, result });
  } catch (e) {
    console.error('[cg/test-email]', e.message);
    res.status(500).json({ success: false, error: e.message });
  }
});

module.exports = { router, setPool };
