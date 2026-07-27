# Fabrika Takip Sistemi — Dash Sürümü (Modül 1)

Bu, mevcut Streamlit uygulamasının **Modül 1'inin** (Excel → Dashboard) Dash ile
yeniden yazılmış hali. Excel'in okunma mantığı (`parser.py`, `validators.py`)
hiç değiştirilmedi — sadece arayüz Streamlit yerine Dash.

## Yerel Test (kendi bilgisayarında)

```bash
python -m venv venv
source venv/bin/activate       # Windows: venv\Scripts\activate
pip install -r requirements.txt

cp .env.example .env
# .env dosyasını aç, SUPABASE_URL ve SUPABASE_KEY'i gir
# (boş bırakırsan otomatik olarak lokal SQLite kullanır)

python app.py
```

Sonra tarayıcıda `http://localhost:8050` adresine git.

## Klasör Yapısı

```
app.py              -> Ana giriş noktası, sayfa yönlendirme (routing)
pages/
  home.py            -> Modül 1: Excel yükle -> Dashboard (TAMAMLANDI)
  module2.py          -> Modül 2: yer tutucu (henüz geliştirilmedi)
  module3.py          -> Modül 3: yer tutucu (henüz geliştirilmedi)
assets/style.css      -> Tüm görsel tasarım burada (Dash bunu otomatik yükler)
parser.py             -> Excel okuma mantığı (DEĞİŞMEDİ)
validators.py          -> Doğrulama mantığı (DEĞİŞMEDİ)
database.py            -> Supabase/SQLite veritabanı katmanı (DEĞİŞMEDİ,
                          sadece Streamlit secrets kaldırıldı, .env kullanılıyor)
calculations.py         -> Repair rate hesaplamaları (DEĞİŞMEDİ)
baseline.py             -> Geçmiş yıl baseline mantığı (DEĞİŞMEDİ)
```

## Render.com'a Deploy Etme

1. Bu kodu kendi GitHub reponuza push edin.
2. [render.com](https://render.com) üzerinde ücretsiz hesap açın, "New Web Service"
   ile reponuzu bağlayın.
3. Ayarlar:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:server`
4. **Environment** sekmesinden `SUPABASE_URL` ve `SUPABASE_KEY` değerlerini girin
   (aynı .env'deki gibi).
5. Deploy edin — birkaç dakika sonra size bir `.onrender.com` linki verecek.

**Not:** Render'ın ücretsiz tier'ı 15 dakika kullanılmayınca uyur, ilk istekte
~30-50 saniye "uyanma" süresi olabilir.

## Henüz Yapılmayanlar (bilerek, Modül 1'e odaklanmak için)

- Pipe-level (boru bazlı) detay analizi
- PDF rapor üretimi
- Baseline (geçmiş yıl) CSV yükleme arayüzü
- Proje gruplama (pipe/machine groups)

Bunlar orijinal Streamlit uygulamasında var, istenirse aynı mantıkla Dash'e
taşınabilir — şimdilik çekirdek akış (Excel → doğrula → kaydet → dashboard)
çalışır durumda olsun diye bunlara girilmedi.
