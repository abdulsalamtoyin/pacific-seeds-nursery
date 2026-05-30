# Sorghum Nursery POC — Level 1

Replaces the VBA macros and paper fieldbook for the **replacements** workflow.

## What's here

```
poc/
  scripts/
    nursery_init.py    # ingest PRISM export → SQLite + workbook + packet PDFs
    fieldbook_export.py # SQLite events → Excel fieldbook for PRISM upload
  backend/
    app.py             # FastAPI sync server (manifest + event sync)
  pwa/
    index.html         # single-screen field tech app
    app.js, sw.js, manifest.webmanifest
  data/
    nursery.sqlite     # created on first run
  output/
    packets.pdf        # printable QR-coded packet labels
    fieldbook.xlsx     # generated workbook
```

## Quick start

```bash
cd poc
source .venv/bin/activate         # python3 -m venv .venv && pip install -r requirements.txt if fresh

# 1. Ingest a PRISM export. Creates SQLite + workbook + QR-coded packet PDFs.
python scripts/nursery_init.py \
    --input "../AUGT1-26S-IMI.xlsx" \
    --nursery-code AUGT1-26S-IMI \
    --sheet Sheet1
```

### Run on your laptop (HTTP, no camera on phones)

```bash
uvicorn backend.app:app --host 0.0.0.0 --port 8000 --reload
# Open http://localhost:8000/         ← camera works (localhost is HTTPS-exempt)
# Open http://localhost:8000/dashboard for the breeder view
```

### Run on your phone in the paddock (HTTPS required for camera)

Browsers refuse to grant camera access over plain `http://<lan-ip>:port`. Generate a
self-signed cert once, then run uvicorn with TLS:

```bash
./scripts/make_cert.sh                       # auto-detects your LAN IP
# or: ./scripts/make_cert.sh 192.168.1.42    # pass IP yourself

.venv/bin/uvicorn backend.app:app --host 0.0.0.0 --port 8443 \
    --ssl-keyfile data/cert.key --ssl-certfile data/cert.crt

# Then on the phone, open:
#   https://<your-mac-ip>:8443/
# Tap "Advanced → Proceed" on the certificate warning (only once per phone).
# After that the camera will work and the PWA can be added to the home screen.
```

## Field tech flow

1. Open the PWA, pick the nursery → device caches the manifest (works offline after).
2. Scan a packet QR → packet details appear.
3. Tap **Replacement** → enter the new entry → Save.
4. Events queue locally; press **Sync** (or it syncs automatically when online).

## When ready to push to PRISM

```bash
python scripts/fieldbook_export.py --nursery-code AUGT1-26S-IMI \
    --out output/fieldbook.xlsx
```

Hand `output/fieldbook.xlsx` (with the populated *Replacements done* tab) to the breeder for PRISM upload.
