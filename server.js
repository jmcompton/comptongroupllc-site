'use strict';
require('dotenv').config();

const express = require('express');
const path    = require('path');
const { Pool } = require('pg');

const app  = express();
const PORT = process.env.PORT || 3001;

// ── Database ──────────────────────────────────────────────────────────────────
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false,
});

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// ── Outreach routes ───────────────────────────────────────────────────────────
const cgRoutes = require('./server/routes/cg-outreach');
cgRoutes.setPool(pool);
app.use('/api/cg', cgRoutes.router);

// ── Contact form ──────────────────────────────────────────────────────────────
app.post('/api/contact', async (req, res) => {
  try {
    const { name, email, subject, message } = req.body;
    if (!name || !email || !message) {
      return res.status(400).json({ success: false, error: 'Missing required fields' });
    }

    // Log to DB if table exists, otherwise just acknowledge
    try {
      await pool.query(
        `INSERT INTO cg_contact_submissions (name, email, subject, message)
         VALUES ($1,$2,$3,$4)`,
        [name, email, subject || 'General', message]
      );
    } catch (_) { /* table may not exist yet — ignore */ }

    // Optionally send via Resend
    try {
      const { Resend } = require('resend');
      const resend = new Resend(process.env.RESEND_API_KEY);
      await resend.emails.send({
        from: 'Compton Group <jmcompton04@gmail.com>',
        to: 'john@comptongroupllc.com',
        subject: `[CGsite] Contact from ${name}: ${subject || 'General'}`,
        text: `Name: ${name}\nEmail: ${email}\nSubject: ${subject || 'General'}\n\n${message}`,
      });
    } catch (_) { /* email optional */ }

    res.json({ success: true, message: 'Message received' });
  } catch (e) {
    console.error('[contact]', e.message);
    res.status(500).json({ success: false, error: 'Server error' });
  }
});

// ── Catch-all: serve homepage ────────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── Bootstrap DB tables + start server ───────────────────────────────────────
app.listen(PORT, async () => {
  console.log(`\n╔══════════════════════════════════════════╗`);
  console.log(`║  Compton Group LLC — Server              ║`);
  console.log(`║  URL:    http://localhost:${PORT}           ║`);
  console.log(`╚══════════════════════════════════════════╝\n`);

  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS cg_contact_submissions (
        id SERIAL PRIMARY KEY,
        name TEXT, email TEXT, subject TEXT, message TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )
    `);
    await pool.query(`
      CREATE TABLE IF NOT EXISTS cg_prospects (
        id SERIAL PRIMARY KEY,
        company TEXT,
        contact_name TEXT,
        contact_title TEXT,
        email TEXT,
        website TEXT,
        industry TEXT,
        revenue_range TEXT,
        ai_opportunity TEXT,
        status TEXT DEFAULT 'pending',
        sequence_type TEXT,
        email1_sent_at TIMESTAMPTZ,
        email2_sent_at TIMESTAMPTZ,
        email3_sent_at TIMESTAMPTZ,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )
    `);
    // Local-targeting + per-prospect signal columns + per-prospect drafted emails.
    // Added via ALTER ... IF NOT EXISTS so existing rows/databases upgrade in place.
    await pool.query(`
      ALTER TABLE cg_prospects
        ADD COLUMN IF NOT EXISTS first_name TEXT,
        ADD COLUMN IF NOT EXISTS city TEXT,
        ADD COLUMN IF NOT EXISTS metro TEXT,
        ADD COLUMN IF NOT EXISTS phone TEXT,
        ADD COLUMN IF NOT EXISTS has_website BOOLEAN,
        ADD COLUMN IF NOT EXISTS online_booking BOOLEAN,
        ADD COLUMN IF NOT EXISTS google_rating NUMERIC(2,1),
        ADD COLUMN IF NOT EXISTS review_count INTEGER,
        ADD COLUMN IF NOT EXISTS category TEXT,
        ADD COLUMN IF NOT EXISTS draft_email1_subject TEXT,
        ADD COLUMN IF NOT EXISTS draft_email1_body TEXT,
        ADD COLUMN IF NOT EXISTS draft_email2_subject TEXT,
        ADD COLUMN IF NOT EXISTS draft_email2_body TEXT,
        ADD COLUMN IF NOT EXISTS draft_email3_subject TEXT,
        ADD COLUMN IF NOT EXISTS draft_email3_body TEXT,
        ADD COLUMN IF NOT EXISTS drafted BOOLEAN DEFAULT FALSE,
        ADD COLUMN IF NOT EXISTS email_type TEXT,
        ADD COLUMN IF NOT EXISTS email_verified BOOLEAN DEFAULT FALSE,
        ADD COLUMN IF NOT EXISTS email_source TEXT
    `);
    await pool.query(`
      CREATE TABLE IF NOT EXISTS cg_sequences (
        id SERIAL PRIMARY KEY,
        sequence_type TEXT UNIQUE,
        email1_subject TEXT DEFAULT '',
        email1_body TEXT DEFAULT '',
        email2_subject TEXT DEFAULT '',
        email2_body TEXT DEFAULT '',
        email3_subject TEXT DEFAULT '',
        email3_body TEXT DEFAULT '',
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )
    `);
    await pool.query(`
      CREATE TABLE IF NOT EXISTS cg_outreach_log (
        id SERIAL PRIMARY KEY,
        prospect_id INTEGER REFERENCES cg_prospects(id) ON DELETE CASCADE,
        sequence_type TEXT,
        email_number INTEGER,
        sent_at TIMESTAMPTZ DEFAULT NOW(),
        resend_id TEXT,
        status TEXT DEFAULT 'sent'
      )
    `);
    // Seed default sequence rows so upsert works
    for (const t of ['efficiency', 'replace_software', 'add_ai']) {
      await pool.query(
        `INSERT INTO cg_sequences (sequence_type) VALUES ($1) ON CONFLICT (sequence_type) DO NOTHING`,
        [t]
      );
    }
    console.log('[DB] Tables ready.');
  } catch (e) {
    console.error('[DB] Setup error:', e.message);
  }
});
