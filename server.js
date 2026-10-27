const http = require("http");
const fs = require("fs");
const path = require("path");

const root = __dirname;

try {
  // Proje klasöründeki .env (çalışma dizini nerede olursa olsun)
  // eslint-disable-next-line global-require
  require("dotenv").config({ path: path.join(root, ".env") });
} catch {
  // dotenv not installed or not needed
}

// Bazı sağlayıcılarda (özellikle app password) boşluklu kopyalanabilir; SMTP için boşluksuz kullanılmalı.
if (process.env.SMTP_PASS != null && process.env.SMTP_PASS !== "") {
  process.env.SMTP_PASS = String(process.env.SMTP_PASS).replace(/\s+/g, "").trim();
}

let nodemailer = null;
try {
  // eslint-disable-next-line global-require
  nodemailer = require("nodemailer");
} catch {
  nodemailer = null;
}

const port = process.env.PORT ? Number(process.env.PORT) : 5500;

function sleep(ms) {
  const n = Number(ms);
  if (!Number.isFinite(n) || n <= 0) return Promise.resolve();
  return new Promise((resolve) => setTimeout(resolve, n));
}

/** Müşteriye otomatik “Mesajınız alındı” maili: varsayılan açık. AUTO_REPLY_ENABLED=false ile kapatılır. */
function autoReplyEnabled() {
  const v = String(process.env.AUTO_REPLY_ENABLED ?? "true")
    .trim()
    .toLowerCase();
  if (v === "0" || v === "false" || v === "no" || v === "off") return false;
  return true;
}

function collectSmtpMissing(env) {
  const smtpMissing = [];
  if (envMissing(env.SMTP_HOST)) smtpMissing.push("SMTP_HOST");
  if (envMissing(env.SMTP_PORT)) smtpMissing.push("SMTP_PORT");
  if (envMissing(env.SMTP_USER)) smtpMissing.push("SMTP_USER");
  if (envMissing(env.SMTP_PASS)) smtpMissing.push("SMTP_PASS");
  if (envMissing(env.MAIL_TO)) smtpMissing.push("MAIL_TO");
  return smtpMissing;
}

function collectGraphMissing(env) {
  const graphMissing = [];
  if (envMissing(env.MS_CLIENT_ID)) graphMissing.push("MS_CLIENT_ID");
  if (envMissing(env.MS_CLIENT_SECRET)) graphMissing.push("MS_CLIENT_SECRET");
  if (envMissing(env.MS_REFRESH_TOKEN)) graphMissing.push("MS_REFRESH_TOKEN");
  if (envMissing(env.MAIL_TO)) graphMissing.push("MAIL_TO");
  return graphMissing;
}

/** auto: Graph tam doluysa Graph; smtp|brevo: her zaman Brevo/SMTP (form bildirimi burada kalsın). */
function shouldTryMicrosoftGraph(env) {
  const mode = String(process.env.MAIL_TRANSPORT || "auto").trim().toLowerCase();
  if (mode === "smtp" || mode === "brevo") return false;
  return collectGraphMissing(env).length === 0;
}

function envMissing(value) {
  if (value == null) return true;
  const t = String(value).trim();
  if (!t) return true;
  const upper = t.toUpperCase();
  if (upper.startsWith("YOUR_")) return true;
  if (upper.includes("YOUR_APP_PASSWORD")) return true;
  // .env.example içindeki yaygın placeholder değerlerini "eksik" kabul et
  if (upper === "BREVO_LOGIN_EMAIL") return true;
  if (upper === "BREVO_SMTP_KEY") return true;
  if (upper === "BREVO_SMTPKEY") return true;
  if (upper === "BREVO_KEY") return true;
  if (/^changeme$/i.test(t)) return true;
  return false;
}

function sendJson(res, statusCode, body) {
  res.writeHead(statusCode, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(body));
}

/** Konu satırı önizlemesi (Outlook liste görünümü için tek satır). */
function subjectSnippet(text, maxLen) {
  const max = Number(maxLen) > 8 ? Number(maxLen) : 72;
  const t = String(text || "")
    .replace(/[\r\n]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!t) return "Yeni mesaj";
  if (t.length <= max) return t;
  return `${t.slice(0, max - 1)}…`;
}

function buildContactPayload({ name, email, message }) {
  const safeName = String(name || "").trim() || "İsimsiz";
  const shortName = safeName.length > 32 ? `${safeName.slice(0, 31)}…` : safeName;
  const snippet = subjectSnippet(message, 64);
  const subject = `[SİTE FORMU] ${shortName} — ${snippet}`;
  const bodyMessage = String(message || "").trim();
  const text = [
    "Acarlar Elektrik - Siteden gelen mesaj",
    "",
    "MESAJ:",
    bodyMessage,
    "",
    "---",
    `Ad Soyad: ${name || "-"}`,
    `E-posta: ${email || "-"}`,
    "",
    `Gönderim: ${new Date().toISOString()}`
  ].join("\n");
  return { subject, text };
}

function buildAutoReplyPayload({ name, message }) {
  const subject = "Mesajınız Alındı | Acarlar Elektrik";
  const safeMessage = String(message || "").trim();
  const intro = [
    `Merhaba ${name || "Değerli Müşterimiz"},`,
    "",
    "Mesajınız tarafımıza ulaştı.",
    "En kısa sürede sizinle iletişime geçeceğiz.",
    "",
    "Acarlar Elektrik"
  ].join("\n");
  const text = safeMessage
    ? ["Gönderdiğiniz mesaj:", safeMessage, "", "---", "", intro].join("\n")
    : intro;
  return { subject, text };
}

function isValidEmail(email) {
  const value = String(email || "").trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

/** Formdaki e-posta iş yeri gelen kutusuyla aynıysa otomatik cevap gönderme (aynı kutuda iki mail, Outlook konuşmada formu gizler). */
function customerEmailIsBusinessInbox(email, mailTo, mailToBackup) {
  const e = String(email || "").trim().toLowerCase();
  if (!e) return false;
  const a = String(mailTo || "").trim().toLowerCase();
  const b = String(mailToBackup || "").trim().toLowerCase();
  if (a && e === a) return true;
  if (b && e === b) return true;
  return false;
}

function shouldSendCustomerAutoReply(email, mailTo, mailToBackup) {
  if (!autoReplyEnabled()) return false;
  if (!isValidEmail(email)) return false;
  if (customerEmailIsBusinessInbox(email, mailTo, mailToBackup)) return false;
  return true;
}

async function sendViaMicrosoftGraph({ name, email, message, env }) {
  const tenantId = (env.MS_TENANT_ID || "consumers").trim();
  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const scope = (env.MS_GRAPH_SCOPE || "https://graph.microsoft.com/Mail.Send offline_access").trim();

  const tokenBody = new URLSearchParams({
    client_id: env.MS_CLIENT_ID.trim(),
    client_secret: env.MS_CLIENT_SECRET.trim(),
    refresh_token: env.MS_REFRESH_TOKEN.trim(),
    grant_type: "refresh_token",
    scope
  });

  const tokenRes = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: tokenBody.toString()
  });

  const tokenRaw = await tokenRes.text();
  let tokenJson = {};
  try {
    tokenJson = tokenRaw ? JSON.parse(tokenRaw) : {};
  } catch {
    tokenJson = {};
  }

  if (!tokenRes.ok || !tokenJson.access_token) {
    const err = new Error("Microsoft token alınamadı.");
    err.code = "EGRAPH_TOKEN";
    err.responseCode = tokenRes.status;
    err.response = tokenRaw || "Token endpoint error";
    throw err;
  }

  const { subject, text } = buildContactPayload({ name, email, message });
  const mailTo = env.MAIL_TO.trim();
  const graphBody = {
    message: {
      subject,
      body: { contentType: "Text", content: text },
      toRecipients: [{ emailAddress: { address: mailTo } }],
      importance: "high"
    },
    saveToSentItems: true
  };

  if (email) {
    graphBody.message.replyTo = [{ emailAddress: { address: email } }];
  }

  const sendRes = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${tokenJson.access_token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(graphBody)
  });

  if (!sendRes.ok) {
    const sendRaw = await sendRes.text();
    const err = new Error("Microsoft Graph ile mail gönderilemedi.");
    err.code = "EGRAPH_SEND";
    err.responseCode = sendRes.status;
    err.response = sendRaw || "Graph send error";
    throw err;
  }

  // Otomatik cevap: sadece ana mail başarıyla gönderildiyse (iş yeri adresiyle test ediliyorsa atlanır)
  let autoReply = null;
  if (shouldSendCustomerAutoReply(email, env.MAIL_TO, env.MAIL_TO_BACKUP)) {
    try {
      // Outlook bazen iki maili aynı anda/sıra karıştırıp otomatik cevabı üstte gösterebiliyor.
      // İş yeri tarafında "önce talep maili" algısı için küçük bir gecikme (AUTO_REPLY_DELAY_MS).
      const delayMs = Number(process.env.AUTO_REPLY_DELAY_MS || "1200");
      await sleep(delayMs);

      const autoPayload = buildAutoReplyPayload({ name, message });
      const autoBody = {
        message: {
          subject: autoPayload.subject,
          body: { contentType: "Text", content: autoPayload.text },
          toRecipients: [{ emailAddress: { address: String(email).trim() } }],
          importance: "normal"
        },
        saveToSentItems: true
      };

      // Müşteri cevap verirse size gelsin
      if (env.MAIL_TO) {
        autoBody.message.replyTo = [{ emailAddress: { address: String(env.MAIL_TO).trim() } }];
      }

      const autoRes = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${tokenJson.access_token}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify(autoBody)
      });

      if (!autoRes.ok) {
        const autoRaw = await autoRes.text();
        autoReply = {
          ok: false,
          code: "EGRAPH_AUTO_REPLY",
          details: autoRaw ? String(autoRaw).slice(0, 300) : undefined
        };
      } else {
        autoReply = { ok: true };
      }
    } catch (err) {
      autoReply = { ok: false, code: err && err.code ? String(err.code) : "AUTO_REPLY_FAILED" };
    }
  } else if (autoReplyEnabled() && isValidEmail(email) && customerEmailIsBusinessInbox(email, env.MAIL_TO, env.MAIL_TO_BACKUP)) {
    autoReply = { ok: true, skipped: "same-as-business-inbox" };
  }

  return { autoReply };
}

async function sendViaSmtp({ name, email, message, env }) {
  const transporter = nodemailer.createTransport({
    host: env.SMTP_HOST,
    port: Number(env.SMTP_PORT),
    secure: String(env.SMTP_SECURE || "true").toLowerCase() === "true",
    requireTLS: true,
    auth: { user: env.SMTP_USER, pass: env.SMTP_PASS },
    tls: { minVersion: "TLSv1.2" }
  });

  const requestedFromEmail = (env.MAIL_FROM_EMAIL || "").trim();
  const smtpUser = String(env.SMTP_USER || "").trim();
  const safeFromEmail = requestedFromEmail || smtpUser;
  const fromName = (env.MAIL_FROM_NAME || "Acarlar Elektrik Site").trim();
  const { subject, text } = buildContactPayload({ name, email, message });
  const primaryTo = String(env.MAIL_TO || "").trim();
  const backupTo = String(env.MAIL_TO_BACKUP || "").trim();
  const recipients = [primaryTo, backupTo]
    .filter(Boolean)
    .filter((value, index, arr) => arr.indexOf(value) === index);

  const mainInfo = await transporter.sendMail({
    envelope: {
      from: safeFromEmail,
      to: recipients
    },
    to: recipients.join(", "),
    from: `${fromName} <${safeFromEmail}>`,
    replyTo: email || env.SMTP_USER,
    subject,
    text,
    headers: {
      "X-Acarlar-Form-Notification": "true",
      "X-Acarlar-Mail-Step": "1-business-notification",
      Importance: "High",
      "X-Priority": "1",
      Priority: "urgent"
    },
    priority: "high"
  });

  let autoReply = null;
  if (shouldSendCustomerAutoReply(email, env.MAIL_TO, env.MAIL_TO_BACKUP)) {
    try {
      const delayMs = Number(process.env.AUTO_REPLY_DELAY_MS || "1200");
      await sleep(delayMs);

      const autoPayload = buildAutoReplyPayload({ name, message });
      const customerTo = String(email).trim();
      const autoInfo = await transporter.sendMail({
        envelope: {
          from: safeFromEmail,
          to: customerTo
        },
        to: customerTo,
        from: `${fromName} <${safeFromEmail}>`,
        replyTo: String(env.MAIL_TO || "").trim() || undefined,
        subject: autoPayload.subject,
        text: autoPayload.text,
        headers: {
          "X-Acarlar-Auto-Reply": "true",
          "X-Acarlar-Mail-Step": "2-customer-auto-reply"
        }
      });
      autoReply = {
        ok: true,
        messageId: autoInfo && autoInfo.messageId ? String(autoInfo.messageId) : undefined
      };
    } catch (err) {
      autoReply = {
        ok: false,
        code: err && err.code ? String(err.code) : "AUTO_REPLY_FAILED"
      };
    }
  } else if (autoReplyEnabled() && isValidEmail(email) && customerEmailIsBusinessInbox(email, env.MAIL_TO, env.MAIL_TO_BACKUP)) {
    autoReply = { ok: true, skipped: "same-as-business-inbox" };
  }

  return { mainInfo, autoReply, recipients };
}

function readJsonBody(req, limitBytes = 40_000) {
  return new Promise((resolve, reject) => {
    let size = 0;
    let raw = "";
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > limitBytes) {
        reject(new Error("Payload too large"));
        req.destroy();
        return;
      }
      raw += chunk.toString("utf8");
    });
    req.on("end", () => {
      try {
        const json = raw ? JSON.parse(raw) : {};
        resolve(json);
      } catch (e) {
        reject(e);
      }
    });
    req.on("error", reject);
  });
}

function contentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  switch (ext) {
    case ".html":
      return "text/html; charset=utf-8";
    case ".css":
      return "text/css; charset=utf-8";
    case ".js":
      return "text/javascript; charset=utf-8";
    case ".json":
      return "application/json; charset=utf-8";
    case ".png":
      return "image/png";
    case ".jpg":
    case ".jpeg":
      return "image/jpeg";
    case ".webp":
      return "image/webp";
    case ".svg":
      return "image/svg+xml";
    case ".ico":
      return "image/x-icon";
    case ".txt":
      return "text/plain; charset=utf-8";
    default:
      return "application/octet-stream";
  }
}

function isImageRequest(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return [".png", ".jpg", ".jpeg", ".webp", ".svg", ".ico"].includes(ext);
}

function sendMissingImagePlaceholder(res, filePath) {
  const name = path.basename(filePath || "image");
  const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="675" viewBox="0 0 1200 675">
  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#0b1220"/>
      <stop offset="100%" stop-color="#111827"/>
    </linearGradient>
  </defs>
  <rect width="1200" height="675" fill="url(#bg)"/>
  <rect x="60" y="60" width="1080" height="555" rx="28" fill="#0f172a" stroke="#fbbf24" stroke-opacity="0.35" stroke-width="4"/>
  <g fill="#e5e7eb" font-family="Arial, sans-serif">
    <text x="110" y="180" font-size="44" font-weight="700">GÖRSEL EKSİK</text>
    <text x="110" y="245" font-size="22" opacity="0.9">Bu dosya repoda yok ya da yolu hatalı:</text>
    <text x="110" y="290" font-size="20" opacity="0.85">${String(name).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</text>
    <text x="110" y="360" font-size="18" opacity="0.75">ÇÖZÜM: İlgili görseli projedeki aynı path ile ekleyip GitHub’a pushlayın.</text>
  </g>
</svg>`;
  res.writeHead(200, { "Content-Type": "image/svg+xml; charset=utf-8", "Cache-Control": "no-store" });
  res.end(svg);
}

function serveFile(filePath, res) {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      if (isImageRequest(filePath)) {
        sendMissingImagePlaceholder(res, filePath);
        return;
      }
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not Found");
      return;
    }
    res.writeHead(200, { "Content-Type": contentType(filePath) });
    res.end(data);
  });
}

function getBaseUrl(req) {
  const headers = req && req.headers ? req.headers : {};
  const host = headers.host ? String(headers.host) : "localhost";
  const xfProto = headers["x-forwarded-proto"] ? String(headers["x-forwarded-proto"]).split(",")[0].trim() : "";
  const proto = xfProto || (req && req.socket && req.socket.encrypted ? "https" : "http");
  return `${proto}://${host}`;
}

function sendText(res, statusCode, text, contentTypeValue = "text/plain; charset=utf-8") {
  res.writeHead(statusCode, {
    "Content-Type": contentTypeValue,
    "Cache-Control": "no-store"
  });
  res.end(text);
}

function buildSitemapXml({ baseUrl }) {
  const urlPaths = [
    "/",
    "/index.html",
    "/projeler.html",
    "/projeler/telefon-tamiri-ve-hat-onarimi.html",
    "/projeler/site-daire-elektrik-tesisati-yenileme.html",
    "/projeler/ofis-aydinlatma-ve-data-altyapisi.html",
    "/projeler/magaza-vitrin-aydinlatma-projesi.html",
    "/projeler/atolye-ve-depo-elektrik-panosu.html",
    "/projeler/guvenlik-kamerasi-ve-internet-altyapisi.html",
    "/projeler/apartman-ortak-alan-aydinlatma.html"
  ];

  const urls = urlPaths
    .map((p) => {
      const loc = `${baseUrl}${p === "/" ? "" : p}`;
      return `  <url><loc>${loc}</loc></url>`;
    })
    .join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${urls}\n</urlset>\n`;
}

const server = http.createServer((req, res) => {
  try {
    const urlPathRaw = (req.url || "/").split("?")[0];

    if (urlPathRaw === "/robots.txt") {
      const baseUrl = getBaseUrl(req);
      const body = `User-agent: *\nAllow: /\n\nSitemap: ${baseUrl}/sitemap.xml\n`;
      sendText(res, 200, body);
      return;
    }

    if (urlPathRaw === "/sitemap.xml") {
      const baseUrl = getBaseUrl(req);
      const xml = buildSitemapXml({ baseUrl });
      sendText(res, 200, xml, "application/xml; charset=utf-8");
      return;
    }

    // API: contact form -> send email
    if (urlPathRaw === "/api/contact") {
      if (req.method === "OPTIONS") {
        res.writeHead(204, {
          "Access-Control-Allow-Origin": "*",
          "Access-Control-Allow-Methods": "POST, OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type"
        });
        res.end();
        return;
      }

      if (req.method !== "POST") {
        sendJson(res, 405, { ok: false, error: "Method not allowed" });
        return;
      }

      readJsonBody(req)
        .then(async (body) => {
          const name = String(body.name || "").trim();
          const email = String(body.email || "").trim();
          const message = String(body.message || "").trim();

          if (!message || message.length < 3) {
            sendJson(res, 400, { ok: false, error: "Mesaj boş olamaz." });
            return;
          }

          const {
            SMTP_HOST,
            SMTP_PORT,
            SMTP_SECURE,
            SMTP_USER,
            SMTP_PASS,
            MAIL_TO,
            MAIL_TO_BACKUP,
            MAIL_FROM_NAME,
            MAIL_FROM_EMAIL,
            MS_TENANT_ID,
            MS_CLIENT_ID,
            MS_CLIENT_SECRET,
            MS_REFRESH_TOKEN,
            MS_GRAPH_SCOPE
          } = process.env;

          const hasGraphConfig = shouldTryMicrosoftGraph({
            MS_CLIENT_ID,
            MS_CLIENT_SECRET,
            MS_REFRESH_TOKEN,
            MAIL_TO
          });

          if (hasGraphConfig) {
            try {
              const graphInfo = await sendViaMicrosoftGraph({
                name,
                email,
                message,
                env: {
                  MS_TENANT_ID,
                  MS_CLIENT_ID,
                  MS_CLIENT_SECRET,
                  MS_REFRESH_TOKEN,
                  MS_GRAPH_SCOPE,
                  MAIL_TO,
                  MAIL_TO_BACKUP
                }
              });
              sendJson(res, 200, { ok: true, via: "microsoft-graph", autoReply: graphInfo ? graphInfo.autoReply : null });
              return;
            } catch (err) {
              const smtpMissingForFallback = collectSmtpMissing({
                SMTP_HOST,
                SMTP_PORT,
                SMTP_USER,
                SMTP_PASS,
                MAIL_TO
              });
              if (nodemailer && smtpMissingForFallback.length === 0) {
                try {
                  const smtpInfo = await sendViaSmtp({
                    name,
                    email,
                    message,
                    env: {
                      SMTP_HOST,
                      SMTP_PORT,
                      SMTP_SECURE,
                      SMTP_USER,
                      SMTP_PASS,
                      MAIL_TO,
                      MAIL_TO_BACKUP,
                      MAIL_FROM_NAME,
                      MAIL_FROM_EMAIL
                    }
                  });
                  sendJson(res, 200, {
                    ok: true,
                    via: "smtp",
                    graphFallback: true,
                    providerResponse:
                      smtpInfo && smtpInfo.mainInfo && smtpInfo.mainInfo.response
                        ? String(smtpInfo.mainInfo.response)
                        : undefined,
                    messageId:
                      smtpInfo && smtpInfo.mainInfo && smtpInfo.mainInfo.messageId
                        ? String(smtpInfo.mainInfo.messageId)
                        : undefined,
                    sentTo: smtpInfo && smtpInfo.recipients ? smtpInfo.recipients : undefined,
                    autoReply: smtpInfo && smtpInfo.autoReply ? smtpInfo.autoReply : undefined
                  });
                  return;
                } catch {
                  // Graph ve SMTP yedek yolu da başarısız; aşağıda Graph hatasını dön
                }
              }
              const code = err && err.code ? String(err.code) : "EGRAPH_UNKNOWN";
              sendJson(res, 502, {
                ok: false,
                code,
                error:
                  "Outlook Graph gönderimi başarısız. Azure App ayarlarını ve refresh token değerini kontrol edin.",
                details: err && err.response ? String(err.response).slice(0, 500) : undefined
              });
              return;
            }
          }

          if (!nodemailer) {
            sendJson(res, 500, { ok: false, error: "Mail modülü eksik. Sunucuya bağımlılıklar kurulmalı." });
            return;
          }

          const smtpMissing = collectSmtpMissing({
            SMTP_HOST,
            SMTP_PORT,
            SMTP_USER,
            SMTP_PASS,
            MAIL_TO
          });

          if (smtpMissing.length) {
            sendJson(res, 500, {
              ok: false,
              missing: smtpMissing,
              error:
                "SMTP ayarları eksik. Proje klasöründe .env dosyası oluşturup şu alanları doldurun: " +
                smtpMissing.join(", ") +
                ". Outlook'ta SMTP AUTH kapalıysa Microsoft Graph alanlarını doldurup OAuth2 ile gönderim kullanın."
            });
            return;
          }

          try {
            const smtpInfo = await sendViaSmtp({
              name,
              email,
              message,
              env: {
                SMTP_HOST,
                SMTP_PORT,
                SMTP_SECURE,
                SMTP_USER,
                SMTP_PASS,
                MAIL_TO,
                MAIL_TO_BACKUP,
                MAIL_FROM_NAME,
                MAIL_FROM_EMAIL
              }
            });

            sendJson(res, 200, {
              ok: true,
              via: "smtp",
              providerResponse:
                smtpInfo && smtpInfo.mainInfo && smtpInfo.mainInfo.response
                  ? String(smtpInfo.mainInfo.response)
                  : undefined,
              messageId:
                smtpInfo && smtpInfo.mainInfo && smtpInfo.mainInfo.messageId
                  ? String(smtpInfo.mainInfo.messageId)
                  : undefined,
              sentTo: smtpInfo && smtpInfo.recipients ? smtpInfo.recipients : undefined,
              autoReply: smtpInfo && smtpInfo.autoReply ? smtpInfo.autoReply : undefined
            });
          } catch (err) {
            const code = err && err.code ? String(err.code) : "UNKNOWN_SMTP_ERROR";
            const rawResponse = err && err.response ? String(err.response) : "";
            const isSmtpAuthDisabled =
              code === "EAUTH" &&
              /basic authentication is disabled/i.test(rawResponse);

            const errorMessage = isSmtpAuthDisabled
              ? "Microsoft hesabında SMTP AUTH kapalı. En kolay çözüm: Brevo SMTP kullanın (smtp-relay.brevo.com) veya Graph OAuth2 alanlarını doldurun."
              : `Mail gönderilemedi. SMTP ayarlarını kontrol edin. (${code})`;

            sendJson(res, 502, {
              ok: false,
              code,
              error: errorMessage
            });
          }
        })
        .catch(() => {
          sendJson(res, 400, { ok: false, error: "Geçersiz istek." });
        });

      return;
    }

    let urlPath = (req.url || "/").split("?")[0];
    if (urlPath === "/" || urlPath === "") urlPath = "/index.html";

    urlPath = decodeURIComponent(urlPath);
    const unsafePath = path.join(root, urlPath);
    const normalized = path.normalize(unsafePath);

    if (!normalized.startsWith(root)) {
      res.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Bad request");
      return;
    }

    let filePath = normalized;
    if (fs.existsSync(filePath) && fs.statSync(filePath).isDirectory()) {
      filePath = path.join(filePath, "index.html");
    }

    serveFile(filePath, res);
  } catch {
    res.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Server error");
  }
});

server.listen(port, "0.0.0.0", () => {
  console.log(`Serving: ${root}`);
  console.log(`URL: http://localhost:${port}/index.html`);
});

