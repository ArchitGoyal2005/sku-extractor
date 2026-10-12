# SKU PDF Extractor Web App — in-memory version

Server-rendered Flask/Jinja app. Uploaded PDFs and generated Excel workbooks are not written to disk.

## Run locally (Windows PowerShell)
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:FLASK_SECRET_KEY = "replace-with-a-long-random-secret"
python app.py
```
Open http://127.0.0.1:5000.

## Docker
```bash
docker build -t sku-pdf-extractor .
docker run --rm -p 8080:8080 -e FLASK_SECRET_KEY="replace-with-a-long-random-secret" sku-pdf-extractor
```

## Configuration
- `FLASK_SECRET_KEY`: set a strong random secret for deployment.
- `MAX_UPLOAD_MB`: maximum PDF upload size, default 50 MB.
- `JOB_TTL_SECONDS`: in-memory results expire after 1800 seconds (30 minutes).
- `MAX_IN_MEMORY_JOBS`: cap on cached jobs, default 20.

## Privacy and deployment
- The upload stream is memory-backed; PDF parsing uses bytes; Excel is generated using `BytesIO`.
- The app does not write the PDF, Excel, result JSON, or page logs to disk. Hosting platform / HTTP server logs may still be persisted by the provider.
- The Excel download endpoint requires a random job token and a browser session owner token; there is no public/static file URL.
- The result is held in RAM until cleared, expired, evicted by the cache limit, or the server restarts. Clearing removes it from the in-process cache.
- Memory use can be significant for large PDFs. Keep a sensible upload limit.
- Run with one worker/process because the cache is process-local. For multiple instances, a shared store would be needed, which is storage (even if not a file).
- Use HTTPS and access controls for public deployments.
- Extraction is tailored to the supplied invoice layout. Image-only PDFs need OCR.
- Logs are displayed after processing finishes, not streamed live.
