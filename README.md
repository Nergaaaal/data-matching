# 🔗 Data Matcher

A minimalist Streamlit application for matching records across multiple CSV/XLSX files.

## Features

- **Multi-file upload** — Upload 2+ CSV or XLSX files
- **Smart column mapping** — Select matching columns via dropdown
- **Manual ID entry** — Add extra IDs to match against uploaded data
- **Full Outer Join** — Case-insensitive, whitespace-stripped matching
- **Styled XLSX export** — Download a professionally formatted match report
- **Match summary** — Visual stats with match rate before downloading

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

The app will open at `http://localhost:8501`.

## Deployment

### Streamlit Community Cloud (Recommended — Free)

1. Push this repo to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io/).
3. Connect your GitHub repo, select `app.py` as the main file.
4. Click **Deploy** — done.

### Docker

```dockerfile
FROM python:3.12-slim
WORKDIR /app
COPY . .
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

Build and run:

```bash
docker build -t data-matcher .
docker run -p 8501:8501 data-matcher
```

> **Note on Vercel:** Streamlit requires a persistent server process, which Vercel's
> serverless architecture does not support. Use Streamlit Community Cloud, Railway,
> Render, or a Docker-based deployment instead.

## Tech Stack

- Python 3.10+
- Streamlit
- pandas
- openpyxl

## License

MIT
