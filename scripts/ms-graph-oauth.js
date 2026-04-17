#!/usr/bin/env node
const path = require("path");

try {
  // eslint-disable-next-line global-require
  require("dotenv").config({ path: path.join(__dirname, "..", ".env") });
} catch {
  // ignore
}

const TENANT = (process.env.MS_TENANT_ID || "consumers").trim();
const CLIENT_ID = (process.env.MS_CLIENT_ID || "").trim();
const CLIENT_SECRET = (process.env.MS_CLIENT_SECRET || "").trim();
const SCOPE = (
  process.env.MS_GRAPH_SCOPE || "https://graph.microsoft.com/Mail.Send offline_access"
).trim();
const REDIRECT_URI = "http://localhost";

function fail(msg) {
  console.error(msg);
  process.exit(1);
}

function buildAuthUrl() {
  if (!CLIENT_ID) {
    fail("MS_CLIENT_ID eksik. .env dosyasını doldurun.");
  }
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    response_type: "code",
    redirect_uri: REDIRECT_URI,
    response_mode: "query",
    scope: SCOPE,
    prompt: "select_account"
  });
  return `https://login.microsoftonline.com/${encodeURIComponent(TENANT)}/oauth2/v2.0/authorize?${params.toString()}`;
}

async function exchangeCode(code) {
  if (!CLIENT_ID) fail("MS_CLIENT_ID eksik.");
  if (!CLIENT_SECRET) fail("MS_CLIENT_SECRET eksik.");
  if (!code) fail("Code parametresi eksik. Örnek: npm run ms:token -- CODE");

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(TENANT)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    code: code.trim(),
    redirect_uri: REDIRECT_URI,
    grant_type: "authorization_code",
    scope: SCOPE
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString()
  });

  const text = await res.text();
  let json = {};
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    json = {};
  }

  if (!res.ok) {
    console.error("Token alma hatası:", res.status);
    console.error(text || "(boş cevap)");
    process.exit(1);
  }

  if (!json.refresh_token) {
    fail("refresh_token dönmedi. İzinleri (Mail.Send + offline_access) kontrol edin.");
  }

  console.log("\nMS_REFRESH_TOKEN:");
  console.log(json.refresh_token);
  console.log("\nBunu .env dosyasına şu şekilde ekleyin:");
  console.log("MS_REFRESH_TOKEN=<yukarıdaki_değer>");
}

async function main() {
  const cmd = process.argv[2];
  if (cmd === "auth-url") {
    console.log(buildAuthUrl());
    return;
  }

  if (cmd === "exchange") {
    await exchangeCode(process.argv[3] || "");
    return;
  }

  console.log("Kullanım:");
  console.log("  npm run ms:auth-url");
  console.log("  npm run ms:token -- CODE");
}

main().catch((err) => {
  console.error(err && err.message ? err.message : String(err));
  process.exit(1);
});

