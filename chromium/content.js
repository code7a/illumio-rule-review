// ----- Illumio Rule Review — Barebones -----
// DRAFT scope + ACTIVE services + scope-aware payload + workloads
// + async submit + poll + auto-disable on zero flows (API match) + sanitized PUT body
// + confirm disable via GET + UI refresh (toolbar click → hash nudge → reload)

(() => {
  'use strict';

  // === Grid & columns (Allow table only) ===================================
  const GRID_SELECTOR           = 'div[data-tid~="comp-grid"][data-tid~="comp-grid-allow"]';
  const ROW_SELECTOR            = 'div[data-tid="comp-grid-row"]';
  const BUTTONS_COL_SELECTOR    = 'div[data-tid="comp-grid-column-buttons"]';
  const EDIT_BUTTON_SELECTOR    = 'button[data-tid~="comp-button"][data-tid~="comp-button-edit"], button[aria-label="Edit"]';

  const COL_TID_CONSUMERS       = 'comp-grid-column-consumers';
  const COL_TID_PROVIDERS       = 'comp-grid-column-providers';
  const COL_TID_SERVICES        = 'comp-grid-column-providingservices';
  const COL_TID_RULENUMBER      = 'comp-grid-column-rulenumber'; // only for query_name cosmetics
  const COL_TID_EXTRASCOPE      = 'comp-grid-column-extrascope'; // "Intra-Scope" / "Extra-Scope"

  // State / diff columns
  const COL_TID_STATE           = 'comp-grid-column-state';
  const DIFF_SIDEBYSIDE_TID     = 'comp-diff-sidebyside';
  const DIFF_ADDED_TID          = 'comp-diff-added';
  const DIFF_REMOVED_TID        = 'comp-diff-removed';
  const COL_TID_DIFFSTATUS      = 'comp-grid-column-diffstatus';  // e.g., “Pending”

  // === Styles ===============================================================
  const CSS = `
    .rr-btn{display:inline-flex;align-items:center;justify-content:center;height:28px;width:28px;border-radius:6px;border:1px solid #d1d5db;background:#fff;color:#111;cursor:pointer;margin-left:6px;padding:0;line-height:1;transition:background .15s ease,border-color .15s ease}
    .rr-btn:hover{background:#f9fafb;border-color:#9ca3af}
    .rr-btn:active{background:#f3f4f6}
    .rr-btn>.rr-ico{display:inline-block;vertical-align:middle}
  `;
  function ensureCssOnce() {
    if (document.getElementById('__rr_min_css__')) return;
    const style = document.createElement('style');
    style.id = '__rr_min_css__';
    style.textContent = CSS;
    document.head?.appendChild(style);
  }
  function hasRRButton(container) {
    return !!container.querySelector('button[data-rr="1"]');
  }

  // === Helpers: text + status ==============================================
  const toLower = (s) => (s || '').toString().trim().toLowerCase();
  const cleanTxt = (el) => (el?.textContent || '').replace(/\u00A0/g, ' ').trim();

  // Returns 'enabled' | 'disabled' | 'unknown'
  function getRowStatus(row) {
    const stateCol = row.querySelector(`[data-tid="${COL_TID_STATE}"]`);
    if (stateCol) {
      const sbs = stateCol.querySelector(`[data-tid="${DIFF_SIDEBYSIDE_TID}"]`);
      if (sbs) {
        const added   = toLower(sbs.querySelector(`[data-tid="${DIFF_ADDED_TID}"]`)?.textContent);
        const removed = toLower(sbs.querySelector(`[data-tid="${DIFF_REMOVED_TID}"]`)?.textContent);
        if (added === 'enabled')  return 'enabled';
        if (added === 'disabled') return 'disabled';
        if (removed === 'enabled')  return 'disabled';
        if (removed === 'disabled') return 'enabled';
      }
      const plain = toLower(stateCol.textContent);
      if (plain.includes('enabled')  && !plain.includes('disabled')) return 'enabled';
      if (plain.includes('disabled') && !plain.includes('enabled'))  return 'disabled';
    }
    const buttonsCell = row.querySelector(BUTTONS_COL_SELECTOR);
    if (buttonsCell) {
      const hasDisable = buttonsCell.querySelector('button[title*="Disable" i], button[aria-label*="Disable" i]');
      const hasEnable  = buttonsCell.querySelector('button[title*="Enable"  i], button[aria-label*="Enable"  i]');
      if (hasDisable) return 'enabled';
      if (hasEnable)  return 'disabled';
    }
    return 'unknown';
  }

  // === Skip ONLY if Diff Status shows "Pending" =============================
  function isPendingUpdate(row) {
    const diffStatusCol = row.querySelector(`[data-tid="${COL_TID_DIFFSTATUS}"]`);
    const diffTxt = toLower(cleanTxt(diffStatusCol));
    return diffTxt.includes('pending');
  }

  // === Determine Intra/Extra scope =========================================
  function getScopeMode(row) {
    const col = row.querySelector(`[data-tid="${COL_TID_EXTRASCOPE}"]`);
    const t = toLower(cleanTxt(col));
    if (t.includes('intra')) return 'intra';
    if (t.includes('extra')) return 'extra';
    return 'unknown';
  }

  // === Column readers -> [{text, href?}] ===================================
  function readPillColumnStructured(row, columnTid) {
    const col = row.querySelector(`[data-tid="${columnTid}"]`);
    if (!col) return [];
    const anchors = Array.from(col.querySelectorAll('a[data-tid^="comp-pill"]'));
    const spans   = Array.from(col.querySelectorAll('span[data-tid^="comp-pill"]'));
    const fromAnchors = anchors.map(a => ({
      text: ((a.querySelector('[data-tid="elem-text"]')?.textContent) || a.textContent || '').trim(),
      href: a.getAttribute('href') || a.href || null
    })).filter(x => x.text);
    const fromSpans = spans.map(s => ({
      text: ((s.querySelector('[data-tid="elem-text"]')?.textContent) || s.textContent || '').trim()
    })).filter(x => x.text);
    const pills = [...fromAnchors, ...fromSpans];
    if (pills.length) return pills;
    const txt = (col.textContent || '').trim();
    return txt ? [{ text: txt }] : [];
  }
  function getRowParts(row) {
    const sources  = readPillColumnStructured(row, COL_TID_CONSUMERS);
    const dests    = readPillColumnStructured(row, COL_TID_PROVIDERS);
    const services = readPillColumnStructured(row, COL_TID_SERVICES);
    return { sources, destinations: dests, services };
  }
  function getRuleNumber(row) {
    const col = row.querySelector(`[data-tid="${COL_TID_RULENUMBER}"]`);
    return (col?.textContent || '').trim() || null;
  }

  // === Ruleset id from URL hash ============================================
  function getRulesetIdFromHash() {
    const m = String(location.hash || '').match(/\/rulesets\/(\d+)\b/);
    return m ? m[1] : null;
  }

  // === OrgId from background ===============================================
  async function getOrgIdFromBackground(retries = 20, delayMs = 250) {
    for (let i = 0; i < retries; i++) {
      try {
        const resp = await chrome.runtime.sendMessage({ type: 'GET_ORG_ID' });
        const id = resp?.orgId ? String(resp.orgId) : null;
        if (id) return id;
      } catch (e) {}
      await new Promise(r => setTimeout(r, delayMs));
    }
    return null;
  }

  // === Service API (ACTIVE) ================================================
  const SERVICE_ID_RE = /\/services\/(\d+)\b/;
  function extractServiceIdFromHref(href) {
    if (!href) return null;
    const m = String(href).match(SERVICE_ID_RE);
    return m ? m[1] : null;
  }

  // Proto‑only preserved: { proto } — no port 0 filler.
  function normalizePort(p) {
    const pick = (...vals) => { for (const v of vals) if (Number.isFinite(v)) return v; return null; };
    const proto = pick(p?.proto, p?.protocol);
    const port  = pick(p?.port, p?.from_port, p?.port_range_start);
    const to    = pick(p?.to_port, p?.port_range_end);
    if (!Number.isFinite(proto)) return null;
    if (!Number.isFinite(port))  return { proto };      // proto-only
    const out = { proto, port };
    if (Number.isFinite(to) && to >= port) out.to_port = to;
    return out;
  }

  async function fetchServicePortsActive(orgId, serviceId) {
    const url = `${location.origin}/api/v2/orgs/${orgId}/sec_policy/active/services/${serviceId}`;
    const res = await fetch(url, { credentials: 'include', headers: { 'Accept': 'application/json' } });
    if (!res.ok) throw new Error(`GET ${url} HTTP ${res.status}`);
    const json = await res.json();
    const rawPorts = Array.isArray(json?.service_ports) ? json.service_ports
                  : Array.isArray(json?.ports)        ? json.ports
                  : [];
    const out = [];
    for (const p of rawPorts) {
      const norm = normalizePort(p);
      if (norm) out.push(norm);
    }
    return out;
  }

  function dedupKey(sp) {
    const proto = Number.isFinite(sp?.proto) ? sp.proto : '';
    const port  = Number.isFinite(sp?.port)  ? sp.port  : '';
    const to    = Number.isFinite(sp?.to_port) ? sp.to_port : '';
    return `${proto}|${port}|${to}`;
  }

  async function buildServicesInclude(servicesPills, orgId) {
    const ids = [...new Set(servicesPills.map(s => extractServiceIdFromHref(s.href)).filter(Boolean))];
    if (!ids.length) return [];
    const seen = new Set();
    const include = [];
    for (const sid of ids) {
      try {
        const ports = await fetchServicePortsActive(orgId, sid);
        for (const it of ports) {
          const key = dedupKey(it);
          if (!seen.has(key)) { seen.add(key); include.push(it); }
        }
      } catch (e) {
        console.warn('[RR] Service fetch failed', { serviceId: sid, error: String(e?.message || e) });
      }
    }
    return include;
  }

  // === Ruleset Scope (DRAFT only) ==========================================
  async function fetchRulesetScopeDraft(orgId, rulesetId) {
    const url = `${location.origin}/api/v2/orgs/${orgId}/sec_policy/draft/rule_sets/${rulesetId}`;
    try {
      const res = await fetch(url, { credentials: 'include', headers: { 'Accept': 'application/json' } });
      if (!res.ok) return null;
      const rs = await res.json();
      const rawScopes = Array.isArray(rs?.scopes) ? rs.scopes : [];
      const clauses = rawScopes
        .map(clause =>
          Array.isArray(clause)
            ? clause.map(s => s?.label?.href).filter(Boolean)
            : []
        )
        .filter(arr => arr.length > 0);
      return { rulesetId, clauses };
    } catch {
      return null;
    }
  }

  // === UI pill → Explorer entity mapping (labels, ip-lists, workloads, ALL) ==
  const UI_RE_LABEL       = /#\/labels\/(\d+)/i;
  const UI_RE_IPLIST      = /#\/iplists\/(\d+)/i;
  const UI_RE_ALLWL       = /#\/workloads(?:$|[/?#])/i;                 // “All Workloads”
  const UI_RE_WORKLOAD_ID = /#\/workloads\/([0-9a-f-]{8}-[0-9a-f-]{4}-[0-9a-f-]{4}-[0-9a-f-]{4}-[0-9a-f-]{12})/i;

  function pillToEntity(pill, orgId) {
    const href = pill.href || '';
    if (UI_RE_WORKLOAD_ID.test(href)) {
      const id = UI_RE_WORKLOAD_ID.exec(href)[1];
      return { workload: { href: `/orgs/${orgId}/workloads/${id}` } };
    }
    if (UI_RE_LABEL.test(href)) {
      const id = UI_RE_LABEL.exec(href)[1];
      return { label: { href: `/orgs/${orgId}/labels/${id}` } };
    }
    if (UI_RE_IPLIST.test(href)) {
      const id = UI_RE_IPLIST.exec(href)[1];
      // DRAFT ip_list path
      return { ip_list: { href: `/orgs/${orgId}/sec_policy/draft/ip_lists/${id}` } };
    }
    if (UI_RE_ALLWL.test(href) || toLower(pill.text) === 'all workloads') {
      return 'ALL';
    }
    return null;
  }

  function scopeLabelsToEntities(scopeLabelHrefs) {
    return (scopeLabelHrefs || []).map(h => ({ label: { href: h } }));
  }

  function applyScopeToSide(baseEntities, scopeLabelEntities, applyScope) {
    if (!applyScope || !scopeLabelEntities.length) return baseEntities;

    const hasALL = baseEntities.includes('ALL');
    let out = baseEntities.filter(e => e !== 'ALL');

    const alreadyHasLabel = out.some(e => e && e.label && e.label.href);

    if (hasALL) {
      out = [...scopeLabelEntities];
    } else if (alreadyHasLabel) {
      const dedupe = new Set(out.map(e => JSON.stringify(e)));
      for (const se of scopeLabelEntities) {
        const key = JSON.stringify(se);
        if (!dedupe.has(key)) { dedupe.add(key); out.push(se); }
      }
    }
    return out;
  }

  function toExplorerIncludeArray(entities) {
    const filtered = entities.filter(e => e !== 'ALL');
    if (filtered.length === 0) return [[]];
    return [filtered];
  }

  // === Time window & payload skeleton ======================================
  function nowIso() { return new Date().toISOString(); }
  function ninetyDaysAgoIso() { return new Date(Date.now() - 90 * 24 * 60 * 60 * 1000).toISOString(); }

  function buildPayloadSkeleton() {
    return {
      sources:      { include: [[]], exclude: [] },
      destinations: { include: [[]], exclude: [] },
      services:     { include: [],  exclude: [] },
      sources_destinations_query_op: 'and',
      start_date: ninetyDaysAgoIso(),
      end_date:   nowIso(),
      policy_decisions: [],
      boundary_decisions: [],
      query_name: 'MAP_QUERY',
      exclude_workloads_from_ip_list_query: true,
      max_results: 1
    };
  }

  // === CSRF helper ==========================================================
  function getCsrfToken() {
    try {
      const meta = document.querySelector('meta[name="csrf-token"]')?.content;
      if (meta) return meta;
      const m = document.cookie.match(/(?:^|;\s*)csrf_token=([^;]+)/);
      if (m) return decodeURIComponent(m[1]);
      const alt = document.cookie.match(/(?:^|;\s*)CSRF-TOKEN=([^;]+)/);
      if (alt) return decodeURIComponent(alt[1]);
    } catch {}
    return null;
  }

  // === Submit async query (create) =========================================
  async function submitAsyncTrafficQuery(orgId, payload) {
    const url = `${location.origin}/api/v2/orgs/${orgId}/traffic_flows/async_queries`;
    const csrf = getCsrfToken();

    const headers = { 'Accept': 'application/json', 'Content-Type': 'application/json' };
    if (csrf) headers['x-csrf-token'] = csrf;

    const res = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers,
      body: JSON.stringify(payload)
    });

    const text = await res.text().catch(() => '');
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch { data = text || null; }

    return { ok: res.ok, status: res.status, data };
  }

  // === Poll async query until terminal =====================================
  function normalizeAsyncHref(href) {
    if (!href) return null;
    if (href.startsWith('http')) return href;
    const path = href.startsWith('/api/') ? href : `/api/v2${href}`;
    return `${location.origin}${path}`;
  }
  function isTerminalStatus(s) {
    const t = String(s || '').toLowerCase();
    return t === 'completed' || t === 'failed' || t === 'canceled' || t === 'timeout';
  }

  async function pollAsyncQuery(orgId, href, opts = {}) {
    const url = normalizeAsyncHref(href);
    if (!url) {
      console.log(JSON.stringify({ async_query_poll_error: 'invalid_href', href }));
      return null;
    }
    const maxWaitMs  = opts.maxWaitMs ?? 5 * 60 * 1000;
    const minDelayMs = opts.minDelayMs ?? 500;
    const maxDelayMs = opts.maxDelayMs ?? 3000;

    let delay = minDelayMs;
    let lastStatus = null;
    const start = Date.now();

    while (Date.now() - start < maxWaitMs) {
      try {
        const res  = await fetch(url, { credentials: 'include', headers: { 'Accept': 'application/json' } });
        const text = await res.text().catch(() => '');
        let data = null;
        try { data = text ? JSON.parse(text) : null; } catch { data = text || null; }

        const status = data?.status ?? data?.state ?? null;
        if (status !== lastStatus) {
          lastStatus = status;
          console.log(JSON.stringify({ async_query_poll: { status, href } }));
        }
        if (!res.ok) {
          console.log(JSON.stringify({ async_query_poll_error: { statusCode: res.status, href, data } }));
          return data;
        }
        if (isTerminalStatus(status)) {
          console.log(JSON.stringify({ async_query_final: data }));
          return data;
        }
      } catch (e) {
        console.log(JSON.stringify({ async_query_poll_exception: String(e?.message || e), href }));
      }

      await new Promise(r => setTimeout(r, delay));
      delay = Math.min(Math.floor(delay * 1.5), maxDelayMs);
    }

    console.log(JSON.stringify({ async_query_timeout: { href, waited_ms: Date.now() - start } }));
    return null;
  }

  // === Draft rules (matching + disable) ====================================
  function normalizeApiHref(href) {
    if (!href) return null;
    if (href.startsWith('http')) return href;
    const path = href.startsWith('/api/') ? href : `/api/v2${href}`;
    return `${location.origin}${path}`;
  }

  async function fetchDraftRules(orgId, rulesetId) {
    const url = `${location.origin}/api/v2/orgs/${orgId}/sec_policy/draft/rule_sets/${rulesetId}/sec_rules`;
    const res = await fetch(url, { credentials: 'include', headers: { 'Accept': 'application/json' } });
    if (!res.ok) {
      console.log(JSON.stringify({ match_rule_lookup: { ok: false, status: res.status } }));
      return [];
    }
    const list = await res.json().catch(() => []);
    console.log(JSON.stringify({ match_rule_lookup: { ok: true, count: Array.isArray(list) ? list.length : 0 } }));
    return Array.isArray(list) ? list : [];
  }

  // helpers to extract IDs from API hrefs
  const RE_LABEL_ID   = /\/labels\/(\d+)\b/;
  const RE_IPLIST_ID  = /\/ip_lists\/(\d+)\b/;
  const RE_WL_UUID    = /\/workloads\/([0-9a-f-]{8}-[0-9a-f-]{4}-[0-9a-f-]{4}-[0-9a-f-]{4}-[0-9a-f-]{12})\b/;
  const RE_SVC_ID     = /\/services\/(\d+)\b/;

  function keyFromEntity(e) {
    if (e === 'ALL') return 'ALL';
    if (e?.label?.href) {
      const m = RE_LABEL_ID.exec(e.label.href); if (m) return `label:${m[1]}`;
    }
    if (e?.ip_list?.href) {
      const m = RE_IPLIST_ID.exec(e.ip_list.href); if (m) return `ip_list:${m[1]}`;
    }
    if (e?.workload?.href) {
      const m = RE_WL_UUID.exec(e.workload.href); if (m) return `workload:${m[1]}`;
    }
    return null;
  }

  function keyFromApiConsumer(c) {
    if (c?.label?.href) {
      const m = RE_LABEL_ID.exec(c.label.href); if (m) return `label:${m[1]}`;
    }
    if (c?.ip_list?.href) {
      const m = RE_IPLIST_ID.exec(c.ip_list.href); if (m) return `ip_list:${m[1]}`;
    }
    if (c?.workload?.href) {
      const m = RE_WL_UUID.exec(c.workload.href); if (m) return `workload:${m[1]}`;
    }
    return null;
  }

  function keyFromApiProvider(p) {
    if (p?.label?.href) {
      const m = RE_LABEL_ID.exec(p.label.href); if (m) return `label:${m[1]}`;
    }
    if (p?.ip_list?.href) {
      const m = RE_IPLIST_ID.exec(p.ip_list.href); if (m) return `ip_list:${m[1]}`;
    }
    if (p?.actors) {
      return 'ALL'; // actors:"ams" => All Workloads
    }
    return null;
  }

  function idsFromApiServices(svcs = []) {
    const out = [];
    for (const s of svcs) {
      const href = s?.href || '';
      const m = RE_SVC_ID.exec(href);
      if (m) out.push(m[1]);
    }
    return out;
  }

  function idsFromServicePills(pills = []) {
    return pills.map(p => extractServiceIdFromHref(p.href)).filter(Boolean);
  }

  function setEq(a, b) {
    if (a.size !== b.size) return false;
    for (const v of a) if (!b.has(v)) return false;
    return true;
  }

  function scoreRuleMatch(rowConsKeys, rowConsAll, rowProvKeys, rowProvAll, rowSvcIds, apiRule) {
    // Build API sides
    const apiConsKeys = new Set((apiRule.consumers || []).map(keyFromApiConsumer).filter(Boolean));
    const apiProvKeys = new Set((apiRule.providers || []).map(keyFromApiProvider).filter(Boolean));
    const apiHasAllConsumers = !!apiRule.unscoped_consumers;
    const apiHasAllProviders = (apiRule.providers || []).some(p => !!p.actors);
    const apiSvcIds = new Set(idsFromApiServices(apiRule.ingress_services || []));

    // base score
    let score = 0;
    let reasons = [];

    // Services: require exact set equality
    const rowSvc = new Set(rowSvcIds);
    if (setEq(rowSvc, apiSvcIds)) { score += 5; reasons.push('svc:eq'); }
    else return { score: -1, reasons: ['svc:ne'] };

    // Consumers
    if (rowConsAll && apiHasAllConsumers) { score += 2; reasons.push('cons:all'); }
    else if (!rowConsAll && !apiHasAllConsumers && setEq(new Set(rowConsKeys), apiConsKeys)) { score += 2; reasons.push('cons:eq'); }
    else if (!rowConsAll && !apiHasAllConsumers) {
      let overlap = 0;
      for (const k of rowConsKeys) if (apiConsKeys.has(k)) overlap++;
      if (overlap > 0) { score += 1; reasons.push('cons:overlap'); }
    }

    // Providers
    if (rowProvAll && apiHasAllProviders) { score += 2; reasons.push('prov:all'); }
    else if (!rowProvAll && !apiHasAllProviders && setEq(new Set(rowProvKeys), apiProvKeys)) { score += 2; reasons.push('prov:eq'); }
    else if (!rowProvAll && !apiHasAllProviders) {
      let overlap = 0;
      for (const k of rowProvKeys) if (apiProvKeys.has(k)) overlap++;
      if (overlap > 0) { score += 1; reasons.push('prov:overlap'); }
    }

    if (apiRule.enabled === true) { score += 0.1; reasons.push('enabled'); }

    return { score, reasons };
  }

  async function matchRuleByRow(orgId, rulesetId, rowSourcePills, rowDestPills, rowServicePills) {
    const rules = await fetchDraftRules(orgId, rulesetId);
    if (!rules.length) return null;

    // Build row key sets from UI pills (BEFORE scope application)
    const rowConsEntities = rowSourcePills.map(p => pillToEntity(p, orgId)).filter(Boolean);
    const rowProvEntities = rowDestPills.map(p => pillToEntity(p, orgId)).filter(Boolean);
    const rowSvcIds = idsFromServicePills(rowServicePills);

    const rowConsAll = rowConsEntities.includes('ALL');
    const rowProvAll = rowProvEntities.includes('ALL');
    const rowConsKeys = rowConsEntities.map(keyFromEntity).filter(k => k && k !== 'ALL');
    const rowProvKeys = rowProvEntities.map(keyFromEntity).filter(k => k && k !== 'ALL');

    let best = null;

    for (const r of rules) {
      const { score, reasons } = scoreRuleMatch(rowConsKeys, rowConsAll, rowProvKeys, rowProvAll, rowSvcIds, r);
      if (score >= 5) {
        if (!best || score > best.score) {
          best = { rule: r, score, reasons };
        }
      }
    }

    if (best) {
      console.log(JSON.stringify({ match_rule_pick: {
        rule_id: best.rule?.id, rule_number: best.rule?.rule_number, score: best.score, reasons: best.reasons
      }}));
      const href = best.rule?.href ||
                   `/orgs/${orgId}/sec_policy/draft/rule_sets/${rulesetId}/sec_rules/${best.rule.id}`;
      return href;
    }

    console.log(JSON.stringify({ match_rule_pick: { reason: 'no_good_match', candidates: rules.length } }));
    return null;
  }

  // --- PUT sanitizer helpers ------------------------------------------------
  function toDraftServiceHref(href, orgId) {
    if (!href) return null;
    const m = RE_SVC_ID.exec(href);
    return m ? `/orgs/${orgId}/sec_policy/draft/services/${m[1]}` : null;
  }

  function stripProviderForPut(p) {
    if (!p || typeof p !== 'object') return null;
    if (p.actors) return { actors: p.actors };
    if (p.label?.href)   return { ...(p.exclusion != null ? { exclusion: !!p.exclusion } : {}), label:   { href: p.label.href } };
    if (p.ip_list?.href) return { ...(p.exclusion != null ? { exclusion: !!p.exclusion } : {}), ip_list: { href: p.ip_list.href } };
    return null;
  }

  function stripConsumerForPut(c) {
    if (!c || typeof c !== 'object') return null;
    const out = {};
    if (c.exclusion != null) out.exclusion = !!c.exclusion;
    if (c.workload?.href) return { ...out, workload: { href: c.workload.href } };
    if (c.label?.href)    return { ...out, label:   { href: c.label.href } };
    if (c.ip_list?.href)  return { ...out, ip_list: { href: c.ip_list.href } };
    return null;
  }

  function sanitizeRuleForPut(rule, orgId) {
    const providers = (rule.providers || []).map(stripProviderForPut).filter(Boolean);
    const consumers = (rule.consumers || []).map(stripConsumerForPut).filter(Boolean);

    const ingress_services = (rule.ingress_services || [])
      .map(s => {
        const href = toDraftServiceHref(s?.href, orgId);
        return href ? { href } : null;
      }).filter(Boolean);

    const egress_services = (rule.egress_services || [])
      .map(s => {
        const href = toDraftServiceHref(s?.href, orgId);
        return href ? { href } : null;
      }).filter(Boolean);

    const body = {
      providers,
      consumers,
      enabled: false,
      ingress_services,
      egress_services,
      network_type: rule.network_type ?? 'brn',
      description: rule.description ?? '',
      consuming_security_principals: Array.isArray(rule.consuming_security_principals) ? rule.consuming_security_principals : [],
      sec_connect: !!rule.sec_connect,
      machine_auth: !!rule.machine_auth,
      stateless: !!rule.stateless,
      unscoped_consumers: !!rule.unscoped_consumers,
      use_workload_subnets: Array.isArray(rule.use_workload_subnets) ? rule.use_workload_subnets : [],
      resolve_labels_as: rule.resolve_labels_as ?? { providers: ['workloads'], consumers: ['workloads'] }
    };

    return body;
  }

  async function getDraftRule(orgId, ruleHref) {
    const url = normalizeApiHref(ruleHref);
    const res = await fetch(url, { credentials: 'include', headers: { 'Accept': 'application/json' } });
    const data = await res.json().catch(() => null);
    console.log(JSON.stringify({ disable_rule_get: { ok: res.ok, status: res.status, href: ruleHref } }));
    return { ok: res.ok, status: res.status, data, url };
  }

  async function putDraftRuleEnabledFalse(ruleUrl, ruleObj, orgId) {
    const csrf = getCsrfToken();
    const headers = { 'Accept': 'application/json', 'Content-Type': 'application/json' };
    if (csrf) headers['x-csrf-token'] = csrf;

    const body = sanitizeRuleForPut(ruleObj, orgId);
    console.log(JSON.stringify({ disable_rule_put_body_preview: body }));

    const res = await fetch(ruleUrl, {
      method: 'PUT',
      credentials: 'include',
      headers,
      body: JSON.stringify(body)
    });

    const text = await res.text().catch(() => '');
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch { data = text || null; }

    console.log(JSON.stringify({ disable_rule_put: { ok: res.ok, status: res.status, href: ruleObj?.href || ruleUrl } }));
    if (!res.ok && res.status === 406) {
      console.log(JSON.stringify({ disable_rule_put_406_hint: 'Send only writable fields; services must be DRAFT hrefs; providers/consumers as actors/label/ip_list/workload with optional exclusion; booleans per rule.' }));
    }
    return { ok: res.ok, status: res.status, data };
  }

  // ✅ NEW: confirm the rule is disabled (GET after PUT)
  async function confirmDisabled(ruleUrl) {
    try {
      const res = await fetch(ruleUrl, { credentials: 'include', headers: { 'Accept': 'application/json' } });
      const json = await res.json().catch(() => null);
      const enabled = !!json?.enabled;
      console.log(JSON.stringify({ disable_rule_confirm: { ok: res.ok, enabled } }));
      return { ok: res.ok, enabled };
    } catch (e) {
      console.log(JSON.stringify({ disable_rule_confirm_error: String(e?.message || e) }));
      return { ok: false, enabled: undefined };
    }
  }

  // ✅ NEW: try to refresh the UI so the disabled state shows up
  function tryRefreshUI() {
    try {
      // 1) Toolbar refresh button (preferred)
      const btn = document.querySelector(
        'button[aria-label*="Refresh" i], [data-tid*="refresh" i], button[title*="Refresh" i]'
      );
      if (btn) {
        btn.click();
        console.log(JSON.stringify({ ui_refresh: 'toolbar_refresh_click' }));
        return;
      }
      // 2) Hash nudge (soft SPA refresh)
      const h = location.hash || '';
      const sep = h.includes('?') ? '&' : '?';
      const temp = `${h}${sep}_rr_refresh=${Date.now()}`;
      location.hash = temp;
      setTimeout(() => { location.hash = h; }, 80);
      console.log(JSON.stringify({ ui_refresh: 'hash_nudge' }));
    } catch {
      // 3) Fallback hard reload
      console.log(JSON.stringify({ ui_refresh: 'reload' }));
      setTimeout(() => location.reload(), 250);
    }
  }

  async function disableRuleByMatchingRow(orgId, rulesetId, rowSourcePills, rowDestPills, rowServicePills) {
    try {
      // 1) Find best-matching rule via API (no UI rule number)
      const ruleHref = await matchRuleByRow(orgId, rulesetId, rowSourcePills, rowDestPills, rowServicePills);
      if (!ruleHref) {
        console.log(JSON.stringify({ disable_rule_lookup: { ok: false, reason: 'no_match' } }));
        return;
      }

      // 2) GET rule JSON
      const got = await getDraftRule(orgId, ruleHref);
      if (!got.ok || !got.data) return;

      // 3) PUT enabled:false with sanitized body
      const putRes = await putDraftRuleEnabledFalse(got.url, got.data, orgId);

      // 4) Confirm & refresh UI
      if (putRes.ok) {
        const conf = await confirmDisabled(got.url);
        if (conf.ok && conf.enabled === false) {
          tryRefreshUI();
        }
      }
    } catch (e) {
      console.log(JSON.stringify({ disable_rule_exception: String(e?.message || e) }));
    }
  }

  // === Button ===============================================================
  function buildButton() {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'rr-btn';
    btn.setAttribute('data-rr', '1');
    btn.setAttribute('title', 'Review');
    btn.setAttribute('aria-label', 'Review');

    const svgNS = 'http://www.w3.org/2000/svg';
    const svg   = document.createElementNS(svgNS, 'svg');
    svg.setAttribute('class', 'rr-ico');
    svg.setAttribute('width', '16');
    svg.setAttribute('height', '16');
    svg.setAttribute('viewBox', '0 0 24 24');
    svg.setAttribute('aria-hidden', 'true');
    svg.setAttribute('focusable', 'false');
    const path  = document.createElementNS(svgNS, 'path');
    path.setAttribute('fill', 'none');
    path.setAttribute('stroke', 'currentColor');
    path.setAttribute('stroke-width', '2');
    path.setAttribute('stroke-linecap', 'round');
    path.setAttribute('stroke-linejoin', 'round');
    path.setAttribute('d', 'M11 5a6 6 0 1 1-4.243 10.243L3.5 18.5M11 7a4 4 0 1 0 0 8 4 4 0 0 0 0-8');
    svg.appendChild(path);
    btn.appendChild(svg);

    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const row = btn.closest(ROW_SELECTOR);
      if (!row) return;

      // Skip ONLY if Pending update
      if (isPendingUpdate(row)) {
        const statusEarly = getRowStatus(row);
        console.log(JSON.stringify({ skipped: true, reason: 'pending_update', status: statusEarly }));
        return;
      }

      // Skip if not enabled
      const status = getRowStatus(row);
      if (status !== 'enabled') {
        console.log(JSON.stringify({ skipped: true, reason: 'not_enabled', status }));
        return;
      }

      // 1) Get org + ruleset id first
      const [orgId, rulesetId] = await Promise.all([
        getOrgIdFromBackground(40, 250),
        Promise.resolve(getRulesetIdFromHash())
      ]);
      if (!orgId) {
        console.log(JSON.stringify({ warning: 'orgId not observed yet from API traffic. Interact with the page and retry.' }));
        return;
      }

      // 2) Fetch DRAFT scope (first)
      const scopeInfo = (rulesetId) ? await fetchRulesetScopeDraft(orgId, rulesetId) : null;
      if (scopeInfo) console.log(JSON.stringify({ ruleset_scope: scopeInfo }));

      // 3) Read row pills + scope mode
      const { sources, destinations, services } = getRowParts(row);
      const scopeMode = getScopeMode(row);
      console.log(JSON.stringify({ status, scopeMode, sources, destinations, services }));

      // 4) Build side entities from original row pills (matching uses raw pills)
      const rowSourcePills  = sources.slice();
      const rowDestPills    = destinations.slice();
      const rowServicePills = services.slice();

      // For payload: apply scope labels
      const sourceBase = rowSourcePills.map(p => pillToEntity(p, orgId)).filter(Boolean);
      const destBase   = rowDestPills.map(p => pillToEntity(p, orgId)).filter(Boolean);

      const scopeLabelHrefs = (scopeInfo?.clauses || []).reduce((acc, clause) => {
        for (const h of clause) acc.add(h);
        return acc;
      }, new Set());
      const scopeLabelEntities = scopeLabelsToEntities([...scopeLabelHrefs]);

      const applyScopeToSrc = (scopeMode === 'intra');
      const applyScopeToDst = (scopeMode === 'intra' || scopeMode === 'extra');

      const sourceFinal = applyScopeToSide(sourceBase, scopeLabelEntities, applyScopeToSrc);
      const destFinal   = applyScopeToSide(destBase,   scopeLabelEntities, applyScopeToDst);

      const sourcesInclude      = toExplorerIncludeArray(sourceFinal);
      const destinationsInclude = toExplorerIncludeArray(destFinal);

      // 5) Services.include from ACTIVE API
      const servicesInclude = await buildServicesInclude(rowServicePills, orgId);

      // 6) Build payload (90 days; max_results=1)
      const serviceNames = rowServicePills.map(s => s.text).filter(Boolean).join(' + ');
      const ruleNo = getRuleNumber(row); // only for query_name cosmetics

      const payload = buildPayloadSkeleton();
      payload.sources.include      = sourcesInclude;
      payload.destinations.include = destinationsInclude;
      payload.services.include     = servicesInclude;
      payload.query_name = serviceNames
        ? `MAP_QUERY_Services Name: ${serviceNames} Time: Last 90 Days`
        : `MAP_QUERY_Row ${ruleNo ?? '?'} Time: Last 90 Days`;

      console.log(JSON.stringify(payload));

      // 7) Submit async query, log create
      const createResp = await submitAsyncTrafficQuery(orgId, payload);
      console.log(JSON.stringify({ async_query_create: createResp }));
      if (!createResp.ok || !createResp?.data?.href) {
        console.log(JSON.stringify({ async_query_create_error: 'missing_href_or_not_ok' }));
        return;
      }

      // 8) Poll to terminal
      const final = await pollAsyncQuery(orgId, createResp.data.href, {
        maxWaitMs: 5 * 60 * 1000,
        minDelayMs: 500,
        maxDelayMs: 3000
      });

      // 9) If completed with zero flows => DISABLE by API MATCH (no UI rule number)
      if (final && String(final.status).toLowerCase() === 'completed') {
        const flowsCount = Number(final?.flows_count ?? final?.matches_count ?? 0);
        if (flowsCount === 0 && rulesetId) {
          console.log(JSON.stringify({ auto_disable_trigger: { reason: 'flows_count_zero' } }));
          await disableRuleByMatchingRow(orgId, rulesetId, rowSourcePills, rowDestPills, rowServicePills);
        }
      }
    });

    return btn;
  }

  // === Injection per row ====================================================
  function addButtonToRow(row) {
    const buttonsCell = row.querySelector(BUTTONS_COL_SELECTOR);
    if (!buttonsCell) return;
    if (hasRRButton(buttonsCell)) return;

    const btn = buildButton();
    const editBtn = buttonsCell.querySelector(EDIT_BUTTON_SELECTOR);
    if (editBtn && editBtn.parentElement) {
      editBtn.insertAdjacentElement('afterend', btn);
    } else {
      buttonsCell.insertBefore(btn, buttonsCell.firstChild);
    }
  }
  function processGrid(grid) {
    grid.querySelectorAll(ROW_SELECTOR).forEach(addButtonToRow);
  }
  function observeGrid(grid) {
    if (grid.__rrObserved) return;
    grid.__rrObserved = true;
    const obs = new MutationObserver((muts) => {
      for (const m of muts) {
        for (const node of m.addedNodes) {
          if (!(node instanceof Element)) continue;
          if (node.matches?.(ROW_SELECTOR)) addButtonToRow(node);
          else node.querySelectorAll?.(ROW_SELECTOR)?.forEach(addButtonToRow);
        }
      }
    });
    obs.observe(grid, { childList: true, subtree: true });
  }
  function ensureAllowGrids() {
    const grids = Array.from(document.querySelectorAll(GRID_SELECTOR));
    if (!grids.length) return false;
    grids.forEach((g) => { processGrid(g); observeGrid(g); });
    return true;
  }

  // === Bootstrap ============================================================
  function init() {
    ensureCssOnce();
    ensureAllowGrids();

    const docObs = new MutationObserver(() => { ensureAllowGrids(); });
    docObs.observe(document.documentElement, { childList: true, subtree: true });

    let attempts = 0;
    const timer = setInterval(() => {
      attempts++;
      if (ensureAllowGrids() || attempts >= 20) clearInterval(timer);
    }, 300);

    window.addEventListener('hashchange', () => setTimeout(ensureAllowGrids, 50));
  }
  if (document.readyState === 'complete' || document.readyState === 'interactive') init();
  else window.addEventListener('DOMContentLoaded', init);
})();