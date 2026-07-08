# Deployment

This project has two parts:

- `src/index.html`, `src/style.css`, and `src/config.js`: static frontend for Netlify.
- `src/main.py`: FastAPI backend that receives the Excel upload and returns the processed file.

Netlify static hosting does not run the FastAPI upload endpoint. Deploy the backend to a Python web host such as Render, Railway, Fly.io, or any server that can run:

```bash
uvicorn src.main:app --host 0.0.0.0 --port $PORT
```

After the backend is deployed, edit `src/config.js`:

```js
window.APP_CONFIG = {
  apiBaseUrl: "https://your-backend-url.example.com",
};
```

For the backend, set `CORS_ORIGINS` to your Netlify site URL, for example:

```text
CORS_ORIGINS=https://silver-faun-7aac4a.netlify.app
```

Leave `apiBaseUrl` empty only when the frontend is served directly by FastAPI.
