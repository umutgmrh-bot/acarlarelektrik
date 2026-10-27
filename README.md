## Siteyi canlıya alma (SEO + domain)

Şu an **domain yok**. Domain aldıktan sonra aşağıdaki yerleri 1 kere güncellemen yeterli.

## İletişim formu mail entegrasyonu

Bu projede iletişim formu, `server.js` üzerinden `/api/contact` endpoint’ine gönderilir.
Gönderim sırası:
1. SMTP (önerilen: Brevo)
2. Microsoft Graph OAuth2 (ileri seviye / opsiyonel)

### Kurulum
- Proje klasöründe:
  - `npm install`
  - `.env.example` dosyasını kopyala → `.env`
  - `.env` içindeki ayarları doldur
  - `npm start`

### En kolay ve güvenilir kurulum (Brevo SMTP)
1) [Brevo](https://www.brevo.com/) hesabı aç.
2) Brevo panelinde `SMTP & API` bölümünden SMTP bilgilerini al:
- SMTP host: `smtp-relay.brevo.com`
- Port: `587`
- Encryption: STARTTLS
- Login: Brevo hesabındaki SMTP kullanıcı bilgisi
- Password: Brevo SMTP key

3) `Senders` bölümünde `acarlarelektrik@outlook.com.tr` adresini sender olarak doğrula (onay mailini onayla).

4) `.env` dosyası örneği:
```env
PORT=5500
SMTP_HOST=smtp-relay.brevo.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=BREVO_LOGIN_EMAIL
SMTP_PASS=BREVO_SMTP_KEY
MAIL_TO=acarlarelektrik@outlook.com.tr
MAIL_FROM_NAME=Acarlar Elektrik Site
MAIL_FROM_EMAIL=acarlarelektrik@outlook.com.tr
```

5) Sunucuyu başlat:
- `npm start`

Bu yöntem Outlook SMTP AUTH engeline takılmaz ve kişisel Outlook hesaplarında en sorunsuz yöntemdir.

### Outlook Graph OAuth2 (opsiyonel, ileri seviye)
Outlook kişisel hesaplarda SMTP basic auth kapalıysa ve Microsoft tarafında App Registration erişimin varsa Graph kullanılabilir.

1) Azure App kaydı oluştur:
- [https://portal.azure.com](https://portal.azure.com) -> App registrations -> New registration
- Accounts in any organizational directory and personal Microsoft accounts seç
- Redirect URI: `http://localhost` (Web)

2) API izni:
- Microsoft Graph -> Delegated permissions -> `Mail.Send`, `offline_access`
- Grant consent (izin onayı)

3) Client secret oluştur:
- Certificates & secrets -> New client secret
- Değeri kopyala

4) Authorization code + refresh token al:
- Komut ile URL üret:
  - `npm run ms:auth-url`
- Çıkan URL'i tarayıcıya yapıştır ve giriş/onay ver
- `http://localhost/?code=...` şeklindeki adresten `code` değerini kopyala
- Refresh token üret:
  - `npm run ms:token -- BURAYA_CODE`
- Komut çıktısındaki token'ı `.env` içindeki `MS_REFRESH_TOKEN` alanına yapıştır

5) `.env` içine yaz:
- `MS_TENANT_ID=consumers`
- `MS_CLIENT_ID=...`
- `MS_CLIENT_SECRET=...`
- `MS_REFRESH_TOKEN=...`
- `MS_GRAPH_SCOPE=https://graph.microsoft.com/Mail.Send offline_access`
Not: Bu alanlar doluysa uygulama otomatik Graph üzerinden gönderir; boşsa SMTP kullanır.

### 1) `sitemap.xml`
- Domain/SSL işleminden sonra ekstra dosya düzenlemenize gerek yok: `server.js` artık `robots.txt` ve `sitemap.xml` çıktısını gelen domain'e göre otomatik üretir.
  - Örnek: `https://acarlarelektrik.com/index.html`

### 2) `robots.txt`
- `Sitemap:` satırı da otomatik olarak doğru domain'e ayarlanır.

### 3) (Önerilen) canonical + og:url
Domain aldıktan sonra `index.html` ve `projeler.html` içine şunlar eklenebilir:
- `<link rel="canonical" href="...">`
- `og:url`

İstersen domaini yazdığında ben bunları da otomatik eklerim.

