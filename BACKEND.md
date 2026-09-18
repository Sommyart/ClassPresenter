# Phase 3 backend

Run `node server.cjs` and open `http://127.0.0.1:3000`. The Node built-in HTTP server serves the existing frontend and exposes JSON endpoints under `/api`.

Data is persisted as JSON in `data/` (created on first run). Set a long random `COOKIE_SECRET` in production; the server refuses to start with the local default when `NODE_ENV=production`. The optional server provides authenticated sessions, owned presentation metadata, presentation sessions/analytics, and a restricted import proxy for public Google Slides and Canva exports. The active product is a focused presentation tool and has no subscription, billing, pricing, or plan-based restrictions.
