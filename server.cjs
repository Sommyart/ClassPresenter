const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const zlib = require("node:zlib");
const { URL } = require("node:url");

const PORT = Number(process.env.PORT) || 3000;
const HOST = process.env.HOST || "127.0.0.1";
const DATA_DIR = path.join(__dirname, "data");
const SESSION_TTL = 1000 * 60 * 60 * 24 * 30;
const COOKIE_NAME = "class_presenter_session";
const COOKIE_SECRET = process.env.COOKIE_SECRET || "change-this-local-secret";
if (process.env.NODE_ENV === "production" && COOKIE_SECRET === "change-this-local-secret") throw new Error("COOKIE_SECRET must be set in production");

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
const files = {
  users: path.join(DATA_DIR, "users.json"),
  sessions: path.join(DATA_DIR, "sessions.json"),
  presentations: path.join(DATA_DIR, "presentations.json"),
  teams: path.join(DATA_DIR, "teams.json"),
  analytics: path.join(DATA_DIR, "analytics.json"),
};
const defaults = { users: [], sessions: [], presentations: [], teams: [], analytics: [] };
for (const [name, file] of Object.entries(files)) {
  if (!fs.existsSync(file)) fs.writeFileSync(file, JSON.stringify(defaults[name], null, 2));
}

const read = (name) => {
  try { return JSON.parse(fs.readFileSync(files[name], "utf8")); } catch { return defaults[name].slice(); }
};
const write = (name, value) => fs.writeFileSync(files[name], JSON.stringify(value, null, 2));
const id = () => crypto.randomUUID();
const now = () => new Date().toISOString();
const safeUser = (user) => ({ id: user.id, name: user.name, email: user.email, organization: user.organization || "", createdAt: user.createdAt });
const json = (res, status, body, headers = {}) => {
  const payload = JSON.stringify(body);
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store", ...headers });
  res.end(payload);
};
const error = (res, status, message) => json(res, status, { error: message });

function hashPassword(password, salt = crypto.randomBytes(16).toString("hex")) {
  return new Promise((resolve, reject) => crypto.scrypt(password, salt, 64, { N: 16384, r: 8, p: 1 }, (err, key) => err ? reject(err) : resolve(`${salt}:${key.toString("hex")}`)));
}
async function verifyPassword(password, encoded) {
  const [salt, expected] = String(encoded).split(":");
  if (!salt || !expected) return false;
  const actual = await hashPassword(password, salt);
  return crypto.timingSafeEqual(Buffer.from(actual.split(":")[1], "hex"), Buffer.from(expected, "hex"));
}
function sign(value) {
  return crypto.createHmac("sha256", COOKIE_SECRET).update(value).digest("base64url");
}
function setSessionCookie(res, sessionId) {
  const value = `${sessionId}.${sign(sessionId)}`;
  res.setHeader("Set-Cookie", `${COOKIE_NAME}=${value}; HttpOnly; SameSite=Lax; Path=/; Max-Age=${SESSION_TTL / 1000}`);
}
function clearSessionCookie(res) {
  res.setHeader("Set-Cookie", `${COOKIE_NAME}=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0`);
}
function cookies(req) {
  return Object.fromEntries((req.headers.cookie || "").split(";").filter(Boolean).map((part) => {
    const index = part.indexOf("="); return [part.slice(0, index).trim(), decodeURIComponent(part.slice(index + 1))];
  }));
}
function currentUser(req) {
  const value = cookies(req)[COOKIE_NAME];
  if (!value) return null;
  const [sessionId, signature] = value.split(".");
  if (!sessionId || !signature) return null;
  const expectedSignature = sign(sessionId);
  if (signature.length !== expectedSignature.length || !crypto.timingSafeEqual(Buffer.from(signature), Buffer.from(expectedSignature))) return null;
  const sessions = read("sessions");
  const session = sessions.find((item) => item.id === sessionId && new Date(item.expiresAt) > new Date());
  if (!session) return null;
  const user = read("users").find((item) => item.id === session.userId);
  return user || null;
}
function requireUser(req, res) {
  const user = currentUser(req);
  if (!user) error(res, 401, "Authentication required");
  return user;
}
function body(req) {
  return new Promise((resolve, reject) => {
    let raw = "";
    req.on("data", (chunk) => { raw += chunk; if (raw.length > 1024 * 1024) req.destroy(); });
    req.on("end", () => { try { resolve(raw ? JSON.parse(raw) : {}); } catch { reject(new Error("Invalid JSON")); } });
    req.on("error", reject);
  });
}
function publicPresentation(item) {
  const { ownerId, ...metadata } = item; return metadata;
}
function ownedPresentation(req, res, presentationId) {
  const user = requireUser(req, res);
  if (!user) return null;
  const item = read("presentations").find((presentation) => presentation.id === presentationId && presentation.ownerId === user.id);
  if (!item) error(res, 404, "Presentation not found");
  return item;
}
async function api(req, res, pathname) {
  try {
    if (req.method === "GET" && pathname === "/api/import-url") {
      const requested = new URL(req.url, `http://${req.headers.host || HOST}`).searchParams.get("url");
      if (!requested) return error(res, 400, "A presentation URL is required");
      const target = new URL(requested);
      const allowed = target.hostname === "docs.google.com" || target.hostname === "canva.com" || target.hostname.endsWith(".canva.com");
      if (!allowed) return error(res, 400, "Only Google Slides and Canva links are supported");
      const upstream = await fetch(target, { redirect: "follow" });
      if (!upstream.ok) return error(res, upstream.status, `The source returned HTTP ${upstream.status}`);
      const contentType = upstream.headers.get("content-type") || "application/octet-stream";
      if (!contentType.includes("pdf") && !contentType.includes("presentation") && !contentType.includes("octet-stream")) return error(res, 415, "The link did not return a downloadable PDF or PPTX");
      const bytes = Buffer.from(await upstream.arrayBuffer());
      if (bytes.length > 100 * 1024 * 1024) return error(res, 413, "Presentation exceeds the 100 MB import limit");
      res.writeHead(200, { "Content-Type": contentType, "Cache-Control": "no-store", "Content-Length": bytes.length });
      return res.end(bytes);
    }
    if (req.method === "POST" && pathname === "/api/auth/signup") {
      const input = await body(req);
      if (!input.email || !input.password || !input.name) return error(res, 400, "Name, email, and password are required");
      if (String(input.password).length < 8) return error(res, 400, "Password must be at least 8 characters");
      const users = read("users");
      const email = String(input.email).trim().toLowerCase();
      if (users.some((item) => item.email === email)) return error(res, 409, "An account with that email already exists");
      const user = { id: id(), name: String(input.name).trim().slice(0, 100), organization: String(input.organization || "").trim().slice(0, 160), email, passwordHash: await hashPassword(String(input.password)), createdAt: now() };
      users.push(user); write("users", users); createSession(user, res); return json(res, 201, safeUser(user));
    }
    if (req.method === "POST" && pathname === "/api/auth/login") {
      const input = await body(req); const user = read("users").find((item) => item.email === String(input.email || "").trim().toLowerCase());
      if (!user || !(await verifyPassword(String(input.password || ""), user.passwordHash))) return error(res, 401, "Invalid email or password");
      createSession(user, res); return json(res, 200, safeUser(user));
    }
    if (req.method === "POST" && pathname === "/api/auth/logout") {
      const value = cookies(req)[COOKIE_NAME]; if (value) { const sessionId = value.split(".")[0]; write("sessions", read("sessions").filter((item) => item.id !== sessionId)); }
      clearSessionCookie(res); return json(res, 200, { ok: true });
    }
    if (req.method === "GET" && pathname === "/api/auth/me") {
      const user = requireUser(req, res); if (!user) return; return json(res, 200, safeUser(user));
    }
    if (req.method === "PATCH" && pathname === "/api/auth/me") {
      const user = requireUser(req, res); if (!user) return;
      const input = await body(req);
      if (input.name !== undefined) user.name = String(input.name).trim().slice(0, 100);
      if (input.organization !== undefined) user.organization = String(input.organization).trim().slice(0, 160);
      const users = read("users"); users[users.findIndex((item) => item.id === user.id)] = user; write("users", users);
      return json(res, 200, safeUser(user));
    }
    if (pathname === "/api/presentations" && req.method === "GET") {
      const user = requireUser(req, res); if (!user) return;
      return json(res, 200, read("presentations").filter((item) => item.ownerId === user.id).map(publicPresentation));
    }
    if (pathname === "/api/presentations" && req.method === "POST") {
      const user = requireUser(req, res); if (!user) return; const input = await body(req);
      const presentations = read("presentations");
      const item = { id: id(), ownerId: user.id, name: String(input.name || "Untitled presentation").slice(0, 200), slideCount: Number(input.slideCount) || 0, createdAt: now(), updatedAt: now(), lastPresented: null };
      presentations.push(item); write("presentations", presentations); return json(res, 201, publicPresentation(item));
    }
    const match = pathname.match(/^\/api\/presentations\/([^/]+)$/);
    if (match) {
      const item = ownedPresentation(req, res, match[1]); if (!item) return;
      if (req.method === "GET") return json(res, 200, publicPresentation(item));
      if (req.method === "PATCH") { const input = await body(req); Object.assign(item, { name: input.name === undefined ? item.name : String(input.name).slice(0, 200), slideCount: input.slideCount === undefined ? item.slideCount : Number(input.slideCount) || 0, updatedAt: now() }); const all = read("presentations"); all[all.findIndex((entry) => entry.id === item.id)] = item; write("presentations", all); return json(res, 200, publicPresentation(item)); }
      if (req.method === "DELETE") { write("presentations", read("presentations").filter((entry) => entry.id !== item.id)); return json(res, 200, { ok: true }); }
    }
    if (req.method === "GET" && pathname === "/api/team") {
      const user = requireUser(req, res); if (!user) return; return json(res, 200, read("teams").filter((team) => team.ownerId === user.id || team.members?.some((member) => member.userId === user.id)));
    }
    if (req.method === "GET" && pathname === "/api/analytics") {
      const user = requireUser(req, res); if (!user) return; return json(res, 200, read("analytics").filter((event) => event.userId === user.id));
    }
    if (req.method === "POST" && pathname === "/api/sessions") {
      const user = requireUser(req, res); if (!user) return; const input = await body(req); const session = { id: id(), userId: user.id, presentationId: input.presentationId || null, startedAt: now(), endedAt: null, durationSeconds: 0, status: "active" }; const sessions = read("sessions"); sessions.push(session); write("sessions", sessions); return json(res, 201, session);
    }
    return error(res, 404, "Not found");
  } catch (err) { return error(res, 400, err.message || "Request failed"); }
}
function createSession(user, res) {
  const session = { id: id(), userId: user.id, createdAt: now(), expiresAt: new Date(Date.now() + SESSION_TTL).toISOString() };
  const sessions = read("sessions").filter((item) => new Date(item.expiresAt) > new Date()); sessions.push(session); write("sessions", sessions); setSessionCookie(res, session.id);
}
const server = http.createServer(async (req, res) => {
  res.setHeader("X-Content-Type-Options", "nosniff"); res.setHeader("X-Frame-Options", "SAMEORIGIN"); res.setHeader("Referrer-Policy", "same-origin");
  const url = new URL(req.url, `http://${req.headers.host || HOST}`);
  if (url.pathname.startsWith("/api/")) return api(req, res, url.pathname);
  if (req.method !== "GET" && req.method !== "HEAD") return error(res, 405, "Method not allowed");
  const requested = url.pathname === "/" ? "/index.html" : url.pathname;
  const file = path.resolve(__dirname, `.${requested}`);
  const root = path.resolve(__dirname);
  if (!(file === root || file.startsWith(`${root}${path.sep}`)) || !fs.existsSync(file) || !fs.statSync(file).isFile()) return error(res, 404, "Not found");
  const types = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8" };
  const contentType = types[path.extname(file)] || "application/octet-stream";
  const acceptsGzip = /\bgzip\b/.test(req.headers["accept-encoding"] || "") && /\.(html|js|css)$/i.test(file);
  res.writeHead(200, { "Content-Type": contentType, "Cache-Control": /\.(js|css)$/i.test(file) ? "public, max-age=3600" : "no-cache", ...(acceptsGzip ? { "Content-Encoding": "gzip", Vary: "Accept-Encoding" } : {}) });
  if (req.method === "HEAD") return res.end();
  const stream = fs.createReadStream(file);
  stream.on("error", () => { if (!res.headersSent) error(res, 500, "Unable to read asset"); else res.destroy(); });
  if (acceptsGzip) stream.pipe(zlib.createGzip()).pipe(res);
  else stream.pipe(res);
});
if (require.main === module) server.listen(PORT, HOST, () => console.log(`ClassPresenter backend listening on http://${HOST}:${PORT}`));
