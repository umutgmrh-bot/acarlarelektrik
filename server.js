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

function envMissing(value) {
  if (value == null) return true;
  const t = String(value).trim();
  if (!t) return true;
  const upper = t.toUpperCase();
  if (upper.startsWith("YOUR_")) return true;
  if (upper.includes("YOUR_APP_PASSWORD")) return true;
  if (/^changeme$/i.test(t)) return true;
  return false;
}

function sendJson(res, statusCode, body) {
  res.writeHead(statusCode, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(body));
}

function buildContactPayload({ name, email, message }) {
  const now = new Date();
  const ticket = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(
    now.getDate()
  ).padStart(2, "0")}-${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}${String(
    now.getSeconds()
  ).padStart(2, "0")}`;
  const subject = `Yeni Form Talebi | ${name || "Adsiz"} | #${ticket}`;
  const text = [
    "Acarlar Elektrik - Yeni Form Talebi",
    "",
    `Talep Kodu: #${ticket}`,
    `Ad Soyad: ${name || "-"}`,
    `E-posta: ${email || "-"}`,
    "",
    "Mesaj:",
    message,
    "",
    `Gönderim zamanı: ${new Date().toISOString()}`
  ].join("\n");
  return { subject, text };
}

function buildAutoReplyPayload({ name }) {
  const subject = "Mesajınız Alındı | Acarlar Elektrik";
  const text = [
    `Merhaba ${name || "Değerli Müşterimiz"},`,
    "",
    "Mesajınız tarafımıza ulaştı.",
    "En kısa sürede sizinle iletişime geçeceğiz.",
    "",
    "Acarlar Elektrik"
  ].join("\n");
  return { subject, text };
}

function isValidEmail(email) {
  const value = String(email || "").trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
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
      toRecipients: [{ emailAddress: { address: mailTo } }]
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
    to: recipients.join(", "),
    from: `${fromName} <${safeFromEmail}>`,
    replyTo: email || env.SMTP_USER,
    subject,
    text,
    headers: {
      "X-Acarlar-Form-Notification": "true",
      Importance: "High",
      "X-Priority": "1",
      Priority: "urgent"
    },
    priority: "high"
  });

  let autoReply = null;
  if (isValidEmail(email)) {
    try {
      const autoPayload = buildAutoReplyPayload({ name });
      const autoInfo = await transporter.sendMail({
        to: String(email).trim(),
        from: `${fromName} <${safeFromEmail}>`,
        replyTo: String(env.MAIL_TO || "").trim() || undefined,
        subject: autoPayload.subject,
        text: autoPayload.text
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

function serveFile(filePath, res) {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not Found");
      return;
    }
    res.writeHead(200, { "Content-Type": contentType(filePath) });
    res.end(data);
  });
}

const server = http.createServer((req, res) => {
  try {
    const urlPathRaw = (req.url || "/").split("?")[0];

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

          const graphMissing = [];
          if (envMissing(MS_CLIENT_ID)) graphMissing.push("MS_CLIENT_ID");
          if (envMissing(MS_CLIENT_SECRET)) graphMissing.push("MS_CLIENT_SECRET");
          if (envMissing(MS_REFRESH_TOKEN)) graphMissing.push("MS_REFRESH_TOKEN");
          if (envMissing(MAIL_TO)) graphMissing.push("MAIL_TO");
          const hasGraphConfig = graphMissing.length === 0;

          if (hasGraphConfig) {
            try {
              await sendViaMicrosoftGraph({
                name,
                email,
                message,
                env: {
                  MS_TENANT_ID,
                  MS_CLIENT_ID,
                  MS_CLIENT_SECRET,
                  MS_REFRESH_TOKEN,
                  MS_GRAPH_SCOPE,
                  MAIL_TO
                }
              });
              sendJson(res, 200, { ok: true, via: "microsoft-graph" });
              return;
            } catch (err) {
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

          const smtpMissing = [];
          if (envMissing(SMTP_HOST)) smtpMissing.push("SMTP_HOST");
          if (envMissing(SMTP_PORT)) smtpMissing.push("SMTP_PORT");
          if (envMissing(SMTP_USER)) smtpMissing.push("SMTP_USER");
          if (envMissing(SMTP_PASS)) smtpMissing.push("SMTP_PASS");
          if (envMissing(MAIL_TO)) smtpMissing.push("MAIL_TO");

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

