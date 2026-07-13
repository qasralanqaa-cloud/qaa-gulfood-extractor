const axios = require('axios');
const https = require('https');
const cheerio = require('cheerio');
const parsePhoneNumber = require('libphonenumber-js').parsePhoneNumberFromString;
const XLSX = require('xlsx');
const fs = require('fs');
const http = require('http');
const path = require('path');

// ============ SIMPLE DOWNLOAD SERVER ============
// Railway needs an HTTP port open to keep the service healthy, and this
// also gives you a real download link for the output file at any time.
const PORT = process.env.PORT || 3000;
http.createServer((req, res) => {
  if (req.url === '/download') {
    const filePath = path.resolve(OUTPUT_FILE);
    if (fs.existsSync(filePath)) {
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="gulfood_2026_full.xlsx"'
      });
      fs.createReadStream(filePath).pipe(res);
    } else {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('File not ready yet. Check /status first.');
    }
  } else if (req.url === '/status') {
    const progress = loadProgress();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      listExtracted: progress.listExtracted,
      totalCompanies: progress.companies ? progress.companies.length : 0,
      enrichedCount: progress.enrichedCount || 0
    }));
  } else {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(`
      <html dir="rtl" lang="fa" style="font-family:Tahoma;text-align:center;padding:50px;">
        <h2>Gulfood Extractor در حال اجراست</h2>
        <p><a href="/status">وضعیت پیشرفت</a></p>
        <p><a href="/download">دانلود فایل اکسل (وقتی آماده باشه)</a></p>
      </html>
    `);
  }
}).listen(PORT, () => console.log(`Download server listening on port ${PORT}`));

const agent = new https.Agent({ rejectUnauthorized: false });
const BASE = 'https://exhibitors.gulfood.com/gulfood-2026/Exhibitor';
const LIMIT = 100;

// If a persistent volume is mounted at /data, use it so progress survives
// restarts/redeploys. Otherwise fall back to the local folder (temporary).
const DATA_DIR = fs.existsSync('/data') ? '/data' : '.';
const PROGRESS_FILE = path.join(DATA_DIR, 'progress.json');
const OUTPUT_FILE = path.join(DATA_DIR, 'gulfood_2026_full.xlsx');
console.log(`Using data directory: ${DATA_DIR} (persistent volume: ${DATA_DIR === '/data'})`);

// ============ PART 1: LIST EXTRACTION ============
async function fetchListPage(start) {
  const body = `limit=${LIMIT}&start=${start}&keyword_search=&cuntryId=&InitialKey=&start_up_exhibitors=&type=&new_category=&new_sub_category=&new_sub_sub_category=&event_sector_value=`;
  const res = await axios.post(`${BASE}/fetchExhibitors`, body, {
    httpsAgent: agent,
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
      'X-Requested-With': 'XMLHttpRequest',
      'Referer': BASE,
      'User-Agent': 'Mozilla/5.0'
    },
    timeout: 30000
  });
  return res.data;
}

function parseListItems(html) {
  const $ = cheerio.load(html);
  const items = [];
  $('.item.list-group-item').each((i, el) => {
    const name = $(el).find('.heading').first().text().trim();
    const standText = $(el).find('.head_discription .web p').first().text().trim();
    const country = $(el).find('.head_discription .web p span').first().text().trim();
    const profileLink = $(el).find('a[href*="/ExbDetails/"]').attr('href') || '';
    const description = $(el).find('.list-group-item-text span').first().text().trim();
    const sectors = $(el).find('.sector_block li').map((j, li) => $(li).text().trim()).get().join('; ');
    const isNumericName = /^\d+$/.test(name.trim());
    if (name) {
      items.push({
        company_name: name,
        name_quality_flag: isNumericName ? 'NUMERIC_CODE_NOT_NAME' : 'ok',
        stand: standText, country, sectors, description, profile_url: profileLink
      });
    }
  });
  return items;
}

async function extractFullList() {
  let all = [];
  let start = 0;
  console.log('=== PHASE 1: Extracting full exhibitor list ===');
  while (true) {
    const html = await fetchListPage(start);
    const items = parseListItems(html);
    if (items.length === 0) break;
    all = all.concat(items);
    console.log(`  start=${start} -> +${items.length} (total: ${all.length})`);
    if (items.length < LIMIT) break;
    start += LIMIT;
    await new Promise(r => setTimeout(r, 400));
  }
  console.log(`Phase 1 done. Total companies: ${all.length}`);
  return all;
}

// ============ PART 2: ENRICHMENT (website -> website, then contact scrape) ============
async function fetchWebsiteFromProfile(profileUrl) {
  try {
    const res = await axios.get(profileUrl, {
      httpsAgent: agent,
      headers: { 'User-Agent': 'Mozilla/5.0' },
      timeout: 15000,
      responseType: 'text',
      transformResponse: [(data) => data]
    });
    const html = typeof res.data === 'string' ? res.data : String(res.data);
    const $ = cheerio.load(html);
    return $('.social_website a').attr('href') || '';
  } catch (e) { return ''; }
}

const CONTACT_PATHS = ['', '/contact', '/contact-us', '/contactus', '/about', '/about-us'];
const EMAIL_REGEX = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
const WHATSAPP_REGEX = /(?:wa\.me\/|api\.whatsapp\.com\/send\?phone=)(\+?\d{7,15})/gi;
const PHONE_REGEX = /(\+?\d[\d\s().-]{7,18}\d)/g;
const JUNK_EMAIL_DOMAINS = ['example.com', 'sentry.io', 'wixpress.com', 'godaddy.com', 'schema.org'];
const PRIORITY_PREFIXES = ['export', 'sales', 'info', 'contact', 'trade', 'business'];

function cleanEmails(raw) {
  const uniq = [...new Set(raw.map(e => e.toLowerCase()))];
  return uniq.filter(e => !JUNK_EMAIL_DOMAINS.some(d => e.endsWith(d)));
}
function pickBestEmail(emails) {
  if (!emails.length) return { email: '', confidence: 0 };
  for (const p of PRIORITY_PREFIXES) {
    const m = emails.find(e => e.startsWith(p + '@') || e.startsWith(p + '.'));
    if (m) return { email: m, confidence: 90 };
  }
  return { email: emails[0], confidence: 60 };
}
function extractPhones(text, countryHint) {
  const raw = text.match(PHONE_REGEX) || [];
  const validated = [];
  for (const r of raw) {
    try {
      const p = parsePhoneNumber(r, countryHint || undefined);
      if (p && p.isValid()) validated.push(p.number);
    } catch (e) {}
  }
  return [...new Set(validated)];
}
function extractWhatsapp(html) {
  const m = [...html.matchAll(WHATSAPP_REGEX)];
  if (!m.length) return '';
  let num = m[0][1];
  if (!num.startsWith('+')) num = '+' + num;
  return num;
}
function normalizeUrl(website, path) {
  try { const u = new URL(website); return `${u.protocol}//${u.host}${path}`; }
  catch (e) { return null; }
}
async function fetchPageSafe(url) {
  try {
    const res = await axios.get(url, {
      httpsAgent: agent,
      headers: { 'User-Agent': 'Mozilla/5.0' },
      timeout: 15000,
      maxRedirects: 5,
      responseType: 'text',
      transformResponse: [(data) => data] // force raw string, never auto-parse JSON
    });
    return typeof res.data === 'string' ? res.data : String(res.data);
  } catch (e) { return null; }
}

async function enrichContact(website, countryHint) {
  const result = { email: '', email_confidence: 0, phones: [], whatsapp: '', whatsapp_confidence: '', status: 'not_attempted' };
  if (!website) { result.status = 'no_website'; return result; }

  let allEmails = [], allPhones = [], whatsapp = '';
  for (const path of CONTACT_PATHS) {
    const url = normalizeUrl(website, path) || (path === '' ? website : null);
    if (!url) continue;
    const html = await fetchPageSafe(url);
    if (!html || typeof html !== 'string') continue;
    const $ = cheerio.load(html);
    allEmails = allEmails.concat(html.match(EMAIL_REGEX) || []);
    allPhones = allPhones.concat(extractPhones($('body').text(), countryHint));
    if (!whatsapp) whatsapp = extractWhatsapp(html);
    if (allEmails.length && allPhones.length && whatsapp) break;
    await new Promise(r => setTimeout(r, 250));
  }

  const cleaned = cleanEmails(allEmails);
  const best = pickBestEmail(cleaned);
  result.email = best.email;
  result.email_confidence = best.confidence;
  result.phones = [...new Set(allPhones)];
  // if no explicit whatsapp link found but we have a phone, mark it as "probable" not confirmed
  if (!whatsapp && result.phones.length) {
    result.whatsapp = result.phones[0];
    result.whatsapp_confidence = 'probable_same_as_phone';
  } else if (whatsapp) {
    result.whatsapp = whatsapp;
    result.whatsapp_confidence = 'confirmed_wa_link';
  }
  result.status = (result.email || result.phones.length) ? 'found' : 'not_found';
  return result;
}

// ============ PART 3: ORCHESTRATION WITH RESUMABLE PROGRESS ============
function loadProgress() {
  if (fs.existsSync(PROGRESS_FILE)) {
    return JSON.parse(fs.readFileSync(PROGRESS_FILE, 'utf-8'));
  }
  return { listExtracted: false, companies: [], enrichedCount: 0 };
}
function saveProgress(progress) {
  fs.writeFileSync(PROGRESS_FILE, JSON.stringify(progress));
}
function saveExcel(companies) {
  const ws = XLSX.utils.json_to_sheet(companies.map(c => ({
    ...c,
    phones: Array.isArray(c.phones) ? c.phones.join(', ') : c.phones
  })));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Gulfood 2026 Full');
  XLSX.writeFile(wb, OUTPUT_FILE);
}

async function main() {
  let progress = loadProgress();

  // Phase 1: extract list (only once)
  if (!progress.listExtracted) {
    const companies = await extractFullList();
    progress.companies = companies;
    progress.listExtracted = true;
    progress.enrichedCount = 0;
    saveProgress(progress);
  } else {
    console.log(`Resuming. List already extracted: ${progress.companies.length} companies.`);
  }

  // Phase 2: get website from profile page (if missing)
  console.log('=== PHASE 2: Fetching website URLs from profiles ===');
  for (let i = 0; i < progress.companies.length; i++) {
    const c = progress.companies[i];
    if (c.website === undefined) {
      c.website = c.profile_url ? await fetchWebsiteFromProfile(c.profile_url) : '';
      if (i % 50 === 0) {
        console.log(`  websites: ${i}/${progress.companies.length}`);
        saveProgress(progress);
      }
      await new Promise(r => setTimeout(r, 200));
    }
  }
  saveProgress(progress);

  // Phase 3: enrich contact info (resumable, skips already-enriched)
  console.log('=== PHASE 3: Enriching email / phone / WhatsApp ===');
  const MEMORY_LIMIT_MB = 350; // restart before Node's heap actually fills up
  for (let i = 0; i < progress.companies.length; i++) {
    const c = progress.companies[i];
    if (c.enriched) continue; // skip already done (resumability)
    let enriched;
    try {
      enriched = await enrichContact(c.website, c.country);
    } catch (e) {
      console.log(`  Skipping ${c.company_name} due to error: ${e.message}`);
      enriched = { email: '', email_confidence: 0, phones: [], whatsapp: '', whatsapp_confidence: '', status: 'error' };
    }
    Object.assign(c, enriched, { enriched: true });
    progress.enrichedCount++;

    if (i % 20 === 0) {
      console.log(`  enriched: ${progress.enrichedCount}/${progress.companies.length}`);
      saveProgress(progress);
      saveExcel(progress.companies); // incremental save - safe to stop/resume anytime

      const heapUsedMB = process.memoryUsage().heapUsed / 1024 / 1024;
      console.log(`  memory: ${heapUsedMB.toFixed(0)} MB`);
      if (heapUsedMB > MEMORY_LIMIT_MB) {
        console.log(`  Memory threshold reached (${heapUsedMB.toFixed(0)}MB). Restarting to free memory. Progress is saved — will resume automatically.`);
        saveProgress(progress);
        saveExcel(progress.companies);
        process.exit(1); // exit "as failure" so Railway's On-Failure restart policy picks it back up
      }
    }
  }

  saveProgress(progress);
  saveExcel(progress.companies);
  console.log(`\n=== DONE. ${progress.companies.length} companies processed. Output: ${OUTPUT_FILE} ===`);
}

main().catch(e => {
  console.error('FATAL:', e.message);
  process.exit(1);
});
