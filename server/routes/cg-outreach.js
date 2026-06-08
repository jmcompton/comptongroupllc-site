'use strict';

// ── CG Outreach Agent Routes ──────────────────────────────────────────────────
// Admin-only (gated client-side by password). No server-side auth needed.
// All routes mount at /api/cg via server.js

const express  = require('express');
const router   = express.Router();
const Anthropic = require('@anthropic-ai/sdk');

// Lazy-init so missing env vars at module-load time don't crash the server
function getAI()     { return new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY }); }

const MODEL      = 'claude-sonnet-4-20250514';
const FROM_EMAIL = process.env.SENDER_EMAIL || 'john@comptongroupllc.com';
const DAILY_LIMIT = 20;

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
router.post('/find-prospects', async (req, res) => {
  const { description } = req.body;
  if (!description || !description.trim()) {
    return res.status(400).json({ error: 'description is required' });
  }

  const userPrompt = `Search the web and find 15 small to mid-size businesses matching this description: ${description.trim()}

For each company return: company name, primary contact name and title, email address, website, industry, estimated revenue range, current software/tech stack if findable, and ONE specific reason why Compton Group LLC (a custom AI software development company) could help them.

Return as a JSON array only, no preamble or markdown. Each element should have these exact keys:
company, contact_name, contact_title, email, website, industry, revenue_range, tech_stack, ai_opportunity`;

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

    // Insert into DB, skip duplicates by email
    let inserted = 0;
    for (const p of prospects) {
      if (!p.company) continue;
      try {
        await pool.query(
          `INSERT INTO cg_prospects
             (company, contact_name, contact_title, email, website, industry, revenue_range, ai_opportunity, status)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8,'pending')
           ON CONFLICT DO NOTHING`,
          [
            p.company || '',
            p.contact_name || '',
            p.contact_title || '',
            p.email || '',
            p.website || '',
            p.industry || '',
            p.revenue_range || '',
            p.ai_opportunity || '',
          ]
        );
        inserted++;
      } catch (rowErr) {
        console.warn('[cg/find-prospects] row insert error:', rowErr.message);
      }
    }

    const all = await pool.query(`SELECT * FROM cg_prospects ORDER BY created_at DESC`);
    res.json({ inserted, prospects: all.rows });
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
// Approve a prospect + auto-select sequence type using AI
router.post('/prospects/:id/approve', async (req, res) => {
  try {
    const r = await pool.query(`SELECT * FROM cg_prospects WHERE id=$1`, [req.params.id]);
    if (!r.rows.length) return res.status(404).json({ error: 'Not found' });
    const prospect = r.rows[0];

    // Quick AI call to pick sequence type
    let sequenceType = 'efficiency'; // default
    try {
      const msg = await getAI().messages.create({
        model: MODEL,
        max_tokens: 50,
        messages: [{
          role: 'user',
          content: `Based on this AI opportunity for a prospect: "${prospect.ai_opportunity}" and their industry: "${prospect.industry}", which outreach angle is best?
Options (reply with ONLY the option key, nothing else):
- efficiency (they waste time on manual processes)
- replace_software (they use outdated/legacy software)
- add_ai (they're doing well but not using AI yet)`,
        }],
      });
      const text = msg.content[0]?.text?.trim().toLowerCase() || '';
      if (text.includes('replace_software') || text.includes('replace')) sequenceType = 'replace_software';
      else if (text.includes('add_ai') || text.includes('add')) sequenceType = 'add_ai';
      else sequenceType = 'efficiency';
    } catch (_) { /* use default */ }

    const updated = await pool.query(
      `UPDATE cg_prospects SET status='approved', sequence_type=$1 WHERE id=$2 RETURNING *`,
      [sequenceType, req.params.id]
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
    // Load sequences
    const seqR = await pool.query(`SELECT * FROM cg_sequences`);
    const seqMap = {};
    for (const s of seqR.rows) seqMap[s.sequence_type] = s;

    // Get approved prospects
    const prospectsR = await pool.query(`
      SELECT * FROM cg_prospects
      WHERE status IN ('approved','sent_1','sent_2')
      LIMIT $1
    `, [DAILY_LIMIT]);

    const prospects = prospectsR.rows;
    if (!prospects.length) {
      return res.json({ sent: 0, message: 'No approved prospects ready for outreach' });
    }

    let sent = 0;
    const errors = [];
    const now = new Date();

    for (const p of prospects) {
      const seq = seqMap[p.sequence_type || 'efficiency'];
      if (!seq) {
        errors.push({ company: p.company, error: 'No sequence configured' });
        continue;
      }

      // Determine which email to send
      let emailNum = null;
      let subject  = null;
      let body     = null;
      let newStatus = null;

      if (p.status === 'approved') {
        emailNum  = 1;
        subject   = seq.email1_subject;
        body      = seq.email1_body;
        newStatus = 'sent_1';
      } else if (p.status === 'sent_1') {
        // Send email 2 only if 5+ days since email 1
        const daysSince = p.email1_sent_at
          ? (now - new Date(p.email1_sent_at)) / 86400000
          : 99;
        if (daysSince < 5) continue;
        emailNum  = 2;
        subject   = seq.email2_subject;
        body      = seq.email2_body;
        newStatus = 'sent_2';
      } else if (p.status === 'sent_2') {
        // Send email 3 only if 12+ days since email 1
        const daysSince = p.email1_sent_at
          ? (now - new Date(p.email1_sent_at)) / 86400000
          : 99;
        if (daysSince < 12) continue;
        emailNum  = 3;
        subject   = seq.email3_subject;
        body      = seq.email3_body;
        newStatus = 'sent_3';
      }

      if (!emailNum || !subject || !body) {
        errors.push({ company: p.company, error: `No content for email ${emailNum}` });
        continue;
      }

      // Personalize subject/body
      const personalSubject = subject
        .replace(/\{company\}/gi, p.company)
        .replace(/\{contact\}/gi, p.contact_name || 'there');
      const personalBody = body
        .replace(/\{company\}/gi, p.company)
        .replace(/\{contact\}/gi, p.contact_name || 'there')
        .replace(/\{industry\}/gi, p.industry || 'your industry');

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
        await sendEmail(p.email || FROM_EMAIL, personalSubject, html); // fallback to self if no email

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
