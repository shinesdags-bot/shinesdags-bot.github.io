// Folia sync: same Supabase account as Neska. Files live at songs/<uid>/folia/<id>__<name>.
const SB_URL = 'https://pxpllllascqigmbyhpmq.supabase.co';
const SB_KEY = 'sb_publishable_T-k5ESUJ2O4i0GjAvHPNZQ_BqellNPc';
const BUCKET = 'songs', SKEY = 'folia.session';
let S = null, busyP = null, refreshing = null, hooks = {};

const errText = (j, f) => { j = j || {}; let m = j.msg || j.error_description || j.message || (typeof j.error === 'string' ? j.error : '') || f || 'Something went wrong.';
  if (/invalid login credentials/i.test(m)) return 'Wrong email or password.';
  if (/already registered|already exists/i.test(m)) return 'That email already has an account. Sign in instead.';
  if (/mime type/i.test(m)) return "The cloud refused this file type.";
  return m.charAt(0).toUpperCase() + m.slice(1); };
const readJSON = r => r.text().then(t => { try { return t ? JSON.parse(t) : null; } catch { return null; } });
async function authPost(path, body) {
  let r; try { r = await fetch(SB_URL + path, { method: 'POST', headers: { apikey: SB_KEY, 'Content-Type': 'application/json' }, body: JSON.stringify(body) }); }
  catch { throw new Error("Can't reach the server. Check your connection."); }
  const j = await readJSON(r); if (!r.ok) throw new Error(errText(j, 'Request failed (' + r.status + ').')); return j || {};
}
function setSession(j) {
  if (!j?.access_token) throw new Error("Sign-in didn't return a session.");
  S = { access: j.access_token, refresh: j.refresh_token, exp: Date.now() + (j.expires_in || 3600) * 1000, user: j.user ? { id: j.user.id, email: j.user.email } : S?.user };
  try { localStorage.setItem(SKEY, JSON.stringify({ refresh: S.refresh, user: S.user })); } catch {}
}
function refresh() {
  if (!S?.refresh) return Promise.reject(new Error('Signed out'));
  return refreshing ||= authPost('/auth/v1/token?grant_type=refresh_token', { refresh_token: S.refresh }).then(setSession).finally(() => { refreshing = null; });
}
async function api(path, { method = 'GET', headers = {}, body, raw } = {}, retried) {
  if (!S) throw new Error('Not signed in');
  if (Date.now() > S.exp - 60000 && !retried) await refresh();
  const h = { apikey: SB_KEY, Authorization: 'Bearer ' + S.access, ...headers };
  if (body != null && !raw) h['Content-Type'] = 'application/json';
  const r = await fetch(SB_URL + path, { method, headers: h, body: body == null ? undefined : raw ? body : JSON.stringify(body) });
  if (r.status === 401 && !retried) { await refresh(); return api(path, { method, headers, body, raw }, true); }
  if (raw === 'blob' && r.ok) return r.arrayBuffer();
  const j = await readJSON(r); if (!r.ok) { const e = new Error(errText(j, 'Request failed (' + r.status + ').')); e.status = r.status; throw e; } return j;
}
const enc = p => p.split('/').map(encodeURIComponent).join('/');
const safe = n => (n || 'Document.pdf').replace(/[\u20a6]/g, 'N').replace(/[\u2013\u2014]/g, '-').replace(/[^A-Za-z0-9 ._()-]/g, '_').slice(0, 120);
const mimeOf = r => /\.docx$/i.test(r.name || '') || r.kind === 'docx' ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 'application/pdf';
const dir = () => S.user.id + '/folia/';
const keyOf = rec => dir() + rec.id + '__' + safe(rec.name);

async function listCloud() {
  const rows = await api('/storage/v1/object/list/' + BUCKET, { method: 'POST', body: { prefix: dir(), limit: 1000, offset: 0, sortBy: { column: 'name', order: 'asc' } } });
  return (rows || []).filter(o => o.id && o.name.includes('__')).map(o => ({ key: dir() + o.name, id: o.name.split('__')[0], name: o.name.slice(o.name.indexOf('__') + 2), ver: o.version || o.metadata?.eTag || o.updated_at, size: o.metadata?.size || 0 }));
}
async function push(rec, cloudKeys) {
  const k = keyOf(rec);
  const res = await api('/storage/v1/object/' + BUCKET + '/' + enc(k), { method: 'POST', headers: { 'Content-Type': mimeOf(rec), 'x-upsert': 'true', 'cache-control': 'no-cache' }, body: new Blob([rec.bytes], { type: mimeOf(rec) }), raw: true });
  for (const o of cloudKeys || []) if (o.id === rec.id && o.key !== k) api('/storage/v1/object/' + BUCKET + '/' + enc(o.key), { method: 'DELETE' }).catch(() => {});
  return res?.version || null;
}
const pull = o => api('/storage/v1/object/authenticated/' + BUCKET + '/' + enc(o.key) + '?v=' + encodeURIComponent(o.ver || ''), { raw: 'blob' }).then(b => new Uint8Array(b));

// write back only if nothing touched the record while we were busy
async function save(rec, patch) { const cur = await hooks.get(rec.id); if (cur && cur.t === rec.t) await hooks.put({ ...cur, ...patch }); }
// Two-way sync. rec: {id,name,bytes,t,synced,cv}. t>synced means local edits not yet uploaded.
async function syncNow() {
  if (!S) return;
  if (busyP) return busyP;
  busyP = (async () => {
    hooks.status?.('syncing');
    const cloud = await listCloud(), byId = new Map(cloud.map(o => [o.id, o]));
    const local = await hooks.all(), seen = new Set();
    for (const rec of local) {
      if (rec.uid && rec.uid !== S.user.id) continue;
      const o = byId.get(rec.id); seen.add(rec.id);
      const changed = !rec.synced || rec.t > rec.synced;
      if (!o) {
        if (rec.synced && !changed) { await hooks.del(rec.id); continue; }   // removed on another device
        { const cv = await push(rec, cloud); await save(rec, { uid: S.user.id, synced: rec.t, cv }); }
      } else if (changed) {
        { const cv = await push(rec, cloud); await save(rec, { uid: S.user.id, synced: rec.t, cv }); }
      } else if (rec.cv == null) {
        await save(rec, { cv: o.ver });                                 // just uploaded; adopt server version
      } else if (o.ver !== rec.cv) {
        const bytes = await pull(o); const now = Date.now();
        await save(rec, { name: o.name, bytes, t: now, synced: now, cv: o.ver });
      }
    }
    for (const o of cloud) if (!seen.has(o.id)) {
      const bytes = await pull(o); const now = Date.now();
      await hooks.put({ id: o.id, name: o.name, kind: /\.docx$/i.test(o.name) ? 'docx' : 'pdf', bytes, t: now, synced: now, cv: o.ver, uid: S.user.id });
    }
    hooks.changed?.();
  })().then(() => hooks.status?.('ok'), e => { console.error(e); hooks.status?.('error', e.message); }).finally(() => { busyP = null; });
  return busyP;
}
async function removeEverywhere(rec) {
  if (!S || !rec.synced) return;
  try { const cloud = await listCloud(); for (const o of cloud) if (o.id === rec.id) await api('/storage/v1/object/' + BUCKET + '/' + enc(o.key), { method: 'DELETE' }); } catch (e) { console.error(e); }
}
async function signIn(email, password, create) {
  const j = create
    ? await authPost('/auth/v1/signup', { email, password }).then(j => j?.access_token ? j : authPost('/auth/v1/token?grant_type=password', { email, password }))
    : await authPost('/auth/v1/token?grant_type=password', { email, password });
  setSession(j); hooks.status?.('ok'); return syncNow();
}
function signOut() { S = null; try { localStorage.removeItem(SKEY); } catch {} hooks.status?.('off'); }
async function restore() {
  let v; try { v = JSON.parse(localStorage.getItem(SKEY) || 'null'); } catch {}
  if (!v?.refresh) { hooks.status?.('off'); return; }
  S = { refresh: v.refresh, user: v.user, exp: 0, access: '' };
  try { await refresh(); syncNow(); }
  catch (e) { if (/connection/i.test(e.message)) hooks.status?.('error', 'Offline'); else signOut(); }
}
export const Sync = { init(h) { hooks = h; restore();
    document.addEventListener('visibilitychange', () => document.visibilityState === 'visible' && syncNow());
    window.addEventListener('online', () => syncNow());
    setInterval(() => document.visibilityState === 'visible' && syncNow(), 60000); },
  signIn, signOut, syncNow, removeEverywhere, get user() { return S?.user || null; } };
