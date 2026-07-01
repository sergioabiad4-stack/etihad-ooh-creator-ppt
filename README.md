# OOH Proposal Generator

A web app for Skyscale that turns an Excel site-data sheet + a PowerPoint template into a fully written, client-ready OOH proposal deck — one slide per site, with AI-generated copy.

---

## How it works

1. User uploads an Excel file (site data) and a PowerPoint template file.
2. The app reads each site row, calls Claude AI to generate taglines, descriptions, and nearby landmarks.
3. It replaces `{PLACEHOLDER}` tokens in the template with real content.
4. A finished `.pptx` file is returned for download.

---

## PowerPoint template — placeholder tokens

Add these tokens as plain text inside your PowerPoint template's text boxes. The app will replace them automatically:

| Token | Replaced with |
|---|---|
| `{SITE_NAME}` | Site Name |
| `{TAGLINE}` | AI-generated punchy tagline |
| `{LOCATION_DESC}` | AI: 2–3 sentence location description |
| `{VISIBILITY_DESC}` | AI: viewing angles & sightlines |
| `{AUDIENCE_DESC}` | AI: audience profile & daily volume |
| `{MARKET}` | Market / city |
| `{LOCATION}` | Location address |
| `{FORMAT}` | Format (e.g. Billboard, Bus Shelter) |
| `{SIZE}` | Physical dimensions |
| `{UNITS}` | Units / Faces |
| `{FREQUENCY}` | Spot Duration + SOV/Loop |
| `{TRAFFIC}` | Daily impacts (comma-formatted) |
| `{LANDMARK_1}` | Nearby landmark 1 with walking time |
| `{LANDMARK_2}` | Nearby landmark 2 with walking time |
| `{LANDMARK_3}` | Nearby landmark 3 with walking time |

Your template should have **one slide**. The app duplicates it for each site row.

---

## Expected Excel columns

```
Market | Site Name | Location | Format | Units/Faces | Size |
Spot Duration | SOV/Loop | Campaign Duration | Impacts |
Net Media Costs | Net Production Costs | Net Total Costs
```

- **Market**: forward-filled automatically (handles merged cells)
- **Site Name**: rows with an empty Site Name are skipped
- **Location = "Various"**: treated as a bus/mobile route — landmarks are replaced with city-wide coverage lines

---

## Installation

```bash
cd ooh-generator
pip install -r requirements.txt
```

---

## Running locally

```bash
# Set your Anthropic API key
export ANTHROPIC_API_KEY=sk-ant-...   # Mac/Linux
set ANTHROPIC_API_KEY=sk-ant-...      # Windows

python app.py
```

Then open http://localhost:5000 in your browser.

---

## Deploying to Render.com (free tier)

1. Push the `ooh-generator` folder to a GitHub repository.

2. Go to [render.com](https://render.com) → **New Web Service** → connect your repo.

3. Configure the service:
   - **Environment**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`

4. Add an environment variable in the Render dashboard:
   - Key: `ANTHROPIC_API_KEY`
   - Value: your key from [console.anthropic.com](https://console.anthropic.com)

5. Click **Deploy**. Render gives you a public URL.

> **Note:** Add `gunicorn` to `requirements.txt` for Render:
> ```
> gunicorn>=21.0.0
> ```
> The free tier spins down after inactivity — first request after sleep may take ~30 seconds.

---

## Project structure

```
ooh-generator/
  app.py              ← Flask backend + processing logic
  requirements.txt
  README.md
  templates/
    index.html        ← Single-page frontend
  uploads/            ← Temp storage for uploaded files (auto-cleaned)
  outputs/            ← Generated .pptx files (auto-cleaned after download)
```
