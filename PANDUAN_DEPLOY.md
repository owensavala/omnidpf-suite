# 🚀 Panduan Deployment OmniPDF (Online & Jaringan)

Aplikasi **OmniPDF** dapat Anda deploy dengan berbagai cara sesuai kebutuhan:

---

## ⚡ Opsi 1: Link Publik Online Instan (Paling Mudah, Gratis & Tanpa Daftar)
Anda sudah memiliki `cloudflared` terpasang di komputer. Anda dapat membagikan web ini ke siapa saja di internet dalam 5 detik:

1. **Klik ganda (double-click) file:**
   👉 **`Buka_Akses_Publik_Online.bat`**
2. Jendela command prompt akan terbuka dan menampilkan link HTTPS gratis, contohnya:
   ```
   https://random-words-xyz.trycloudflare.com
   ```
3. Salin link tersebut dan buka di HP, tablet, atau kirimkan ke rekan Anda. Web sudah online dan dapat diakses dari mana saja di seluruh dunia selama komputer Anda menyala!

---

## 📶 Opsi 2: Akses Jaringan Lokal (Satu WiFi / Kantor / Rumah)
Jika Anda hanya ingin mengakses aplikasi dari HP, tablet, atau komputer lain yang terhubung ke WiFi yang sama:

1. Jalankan aplikasi seperti biasa:
   👉 **`Buka_iLovePDF_Offline.bat`**
2. Cari tahu IP lokal laptop Anda (misalnya IP Anda adalah `10.154.226.65` atau `192.168.1.xxx`).
3. Buka browser di HP/perangkat lain yang satu WiFi lalu ketik:
   ```
   http://10.154.226.65:5000
   ```
4. Aplikasi langsung terbuka dan siap digunakan tanpa menggunakan kuota internet!

---

## ☁️ Opsi 3: Deploy Online 24/7 Gratis ke Cloud (Render / Railway / Hugging Face)

File konfigurasi untuk deployment cloud sudah disiapkan lengkap:
- `Procfile` (untuk Render / Railway / Heroku)
- `render.yaml` (Blueprint Render)
- `requirements.txt` (Daftar library)

### Langkah Deploy ke Render (Gratis 24/7):
1. Buat akun di [Render.com](https://render.com) (bisa login via GitHub).
2. Upload folder ini ke repository GitHub Anda (Public atau Private).
3. Di dashboard Render, klik **New +** $\rightarrow$ **Web Service**.
4. Pilih repository GitHub Anda.
5. Render akan mendeteksi konfigurasi secara otomatis:
   - **Environment:** `Python`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn -w 2 --threads 4 -b 0.0.0.0:$PORT server:app --timeout 180`
6. Klik **Create Web Service**. Dalam 2 menit aplikasi Anda akan aktif dengan domain permanen seperti `https://omnidpf-xyz.onrender.com`.

---

## 🐳 Opsi 4: Deploy dengan Docker (VPS / Linux / Server Sendiri)

Jika Anda memiliki VPS (DigitalOcean, AWS, IDCloudHost, Hetzner, dll.) atau ingin menjalankannya di Docker desktop:

Cukup jalankan satu perintah:
```bash
docker compose up -d --build
```
Aplikasi akan otomatis dibungkus ke dalam container terisolasi yang stabil dan berjalan di port `5000`.
