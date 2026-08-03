# Factory Tracking System — Dash Version

A Dash rewrite of Module 1 (Excel → Dashboard) of the factory daily activity
tracking tool. The Excel-reading logic (`parser.py`, `validators.py`) uses
the same fixed cell/column contract as the original spec — only the UI is
built with Dash instead of Streamlit.

## Local Testing (on your own machine)

```bash
python -m venv venv
source venv/bin/activate       # Windows: venv\Scripts\activate
pip install -r requirements.txt

cp .env.example .env
# Open .env and fill in SUPABASE_URL and SUPABASE_KEY
# (leave them empty to automatically use local SQLite)

python app.py
```

Then open `http://localhost:8050` in your browser.

## Folder Structure

```
app.py              -> Entry point, page routing
pages/
  home.py            -> Module 1: Excel upload -> Dashboard (procesing)
  module2.py          -> Module 2: placeholder (not yet built)
  module3.py          -> Module 3: placeholder (not yet built)
assets/style.css      -> All visual styling lives here (Dash auto-loads it)
parser.py             -> Excel-reading logic
validators.py          -> Validation logic
database.py            -> Supabase/SQLite database layer
calculations.py         -> Repair rate calculations
baseline.py             -> Historical baseline logic
supabase_setup.sql      -> Supabase schema (run once in the SQL Editor)
```

## Database Setup (Supabase)

The app uses Supabase in production so data survives server restarts and
redeploys. If `SUPABASE_URL` / `SUPABASE_KEY` are not set, it automatically
falls back to a local SQLite file — handy for quick local testing, but data
does not persist across deploys.

1. Create a project at [supabase.com](https://supabase.com).
2. Open **SQL Editor** and run the full contents of `supabase_setup.sql` once.
3. Go to **Project Settings > API** and copy the **Project URL** and the
   **service_role** secret key (not the anon/public key — the backend needs
   to bypass Row Level Security).
4. Put both values into `.env`:
   ```
   SUPABASE_URL=https://your-project-ref.supabase.co
   SUPABASE_KEY=your-service-role-key
   ```

## Deploying to Render.com

1. Push this code to your own GitHub repo.
2. Create a free account at [render.com](https://render.com) and connect a
   "New Web Service" to your repo.
3. Settings:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:server`
4. In the **Environment** tab, set `SUPABASE_URL` and `SUPABASE_KEY` (same
   values as in `.env`).
5. Deploy — after a few minutes you'll get a `.onrender.com` link.

**Note:** Render's free tier sleeps after 15 minutes of inactivity, so the
first request afterward may take ~30-50 seconds to wake up.

## Not Yet Built

- Pipe-level detail analysis
- PDF report generation
- Project grouping (pipe/machine groups)

These can be added to Dash later, using the same underlying data model that
already exists in `database.py` — the core flow (Excel → validate → save →
dashboard) works today.
