// illumio-rule-review.js

(function () {
  'use strict';

  /***************************************************************************
   * SETTINGS (defaults + persisted)
   ***************************************************************************/
  const SETTINGS_KEY = 'illumio-rulecore-settings-v1';
  const DEFAULTS = {
    lookbackDays: 90,
    pollIntervalMs: 2500,
    maxPollMs: 5 * 60 * 1000,
    maxResults: 200000,

    // Zero-flow handling (batch modes honor this; single-run shows manual Disable)
    autoDisableOnZeroFlows: true,
    reloadAfterDisable: true, // Ruleset batch reload is deferred to end; Policies reloads only if changes applied; single-run reloads after you click Disable (if enabled)

    // Tightening (flows > 0)
    tightenIpListsEnabled: true,

    // Logging / HUD
    verbose: true,
    verbosePoll: true,
    logRawJson: true,
    showToasts: true,
    hudEnabled: true
  };

  function loadSettings(def) {
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      return raw ? { ...def, ...JSON.parse(raw) } : { ...def };
    } catch { return { ...def }; }
  }
  function saveSettings() {
    try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(SETTINGS)); } catch {}
  }
  const SETTINGS = loadSettings(DEFAULTS);

  /***************************************************************************
   * LOGGING + TOAST
   ***************************************************************************/
  const log   = (...a) => SETTINGS.verbose && console.log('[RuleCore]', ...a);
  const warn  = (...a) => console.warn('[RuleCore]', ...a);
  const error = (...a) => console.error('[RuleCore]', ...a);
  function logRaw(label, data) {
    if (!SETTINGS.verbose) return;
    if (SETTINGS.logRawJson) {
      if (typeof data === 'string') console.log(`[RuleCore][RAW] ${label}:\n${data}`);
      else { try { console.log(`[RuleCore][RAW] ${label}:\n${JSON.stringify(data, null, 2)}`); } catch { console.log('[RuleCore][RAW] (unstringifiable)', data); } }
    } else { console.log('[RuleCore]', label, data); }
  }
  function toast(msg, type = 'info', ms = 3200) {
    if (!SETTINGS.showToasts) return;
    const id = 'rulecore-toast-container';
    let c = document.getElementById(id);
    if (!c) {
      c = document.createElement('div'); c.id = id;
      c.style.cssText = `position:fixed;right:16px;bottom:16px;z-index:999999;display:flex;flex-direction:column;gap:8px;font:12px/1.4 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;`;
      if (!document.body) { window.addEventListener('DOMContentLoaded', () => document.body.appendChild(c)); }
      else document.body.appendChild(c);
    }
    const el = document.createElement('div');
    el.textContent = String(msg || '');
    el.style.cssText = `padding:8px 10px;border-radius:8px;color:#111;background:#fff;border:1px solid #e5e7eb;box-shadow:0 10px 24px rgba(0,0,0,.15);max-width:560px;`;
    if (type === 'ok') el.style.borderColor = '#10b981';
    if (type === 'warn') el.style.borderColor = '#f59e0b';
    if (type === 'err') el.style.borderColor = '#ef4444';
    c.appendChild(el); setTimeout(() => el.remove(), ms);
  }

  /***************************************************************************
   * HUDs — Single, Ruleset Batch, Policies Batch
   ***************************************************************************/
  let HUD=null, HUD_STATUS=null, HUD_RULEID=null, HUD_FLOWS=null, HUD_BAR=null;
  let HUD_DECISION=null, HUD_BTN_DISABLE=null;
  let HUD_SETTINGS_WRAP=null, HUD_GEAR_BTN=null;
  let HUD_TIGHTEN=null, HUD_TIGHTEN_APPLY=null, HUD_TIGHTEN_SKIP=null;

  // Ruleset Batch HUD
  let BHUD=null, BHUD_STATUS=null, BHUD_COUNTERS=null, BHUD_BAR=null, BHUD_CURR=null, BHUD_DETAILS=null;
  let BATCH_ABORT_REQUESTED=false;
  let BATCH_MODE=false;

  // Policies Batch HUD
  let PHUD=null, PHUD_STATUS=null, PHUD_BAR=null, PHUD_COUNTERS_RS=null, PHUD_COUNTERS_RULES=null, PHUD_CURR_RS=null, PHUD_CURR_RULE=null, PHUD_DETAILS=null;
  let POLICIES_ABORT_REQUESTED=false;

  let ABORT_REQUESTED=false; // For single-rule runs only

  const STEPS = {
    start:{pct:5,label:'Starting…'},
    fetchRuleset:{pct:10,label:'Fetching ruleset…'},
    matchRule:{pct:25,label:'Matching rule…'},
    buildSignature:{pct:40,label:'Applying scope…'},
    createQuery:{pct:55,label:'Creating query…'},
    pollingStart:{pct:70,label:'Waiting for results…'},
    pollingTick:{pct:70,label:'Waiting for results…'},
    complete:{pct:100,label:'Completed'},
    cancelled:{pct:100,label:'Cancelled'},
    skipped:{pct:100,label:'Skipped'},
    error:{pct:100,label:'Error'}
  };

  /** Single-rule HUD  */
  function createHUD(ruleId='—') {
    if (!SETTINGS.hudEnabled) return;
    removeHUD();
    const el = document.createElement('div');
    el.id='illumio-rulecore-hud';
    el.style.cssText = `position:fixed;left:16px;bottom:16px;width:520px;background:#fff;color:#111;z-index:100000;border:1px solid #e5e7eb;border-radius:12px;box-shadow:0 14px 30px rgba(0,0,0,0.18);font:13px/1.4 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;overflow:hidden;`;
    const disableBtnText = SETTINGS.reloadAfterDisable ? 'Disable & Reload' : 'Disable';
    el.innerHTML = `
      <div style="padding:12px 14px;border-bottom:1px solid #eee;display:flex;align-items:center;justify-content:space-between;">
        <div style="display:flex;align-items:center;gap:10px;">
          <div style="font-weight:700;">Rule Review</div>
          <button id="hud-gear" title="Settings" style="border:1px solid #d1d5db;background:#fff;border-radius:50%;width:28px;height:28px;cursor:pointer;display:flex;align-items:center;justify-content:center;">⚙️</button>
        </div>
        <div style="display:flex;gap:8px;">
          <button id="hud-cancel" style="border:1px solid #ef4444;background:#fff;color:#ef4444;border-radius:8px;padding:4px 10px;cursor:pointer;">Cancel</button>
          <button id="hud-close"  style="border:1px solid #d1d5db;background:#fff;border-radius:8px;padding:4px 10px;cursor:pointer;">Close</button>
        </div>
      </div>
      <div style="padding:10px 14px;display:flex;align-items:center;justify-content:space-between;">
        <div style="font-size:12px;color:#555;">Rule ID: <strong id="hud-ruleid">${escapeHtml(ruleId)}</strong></div>
        <div style="font-size:12px;color:#444;">Flows: <strong id="hud-flows">—</strong></div>
      </div>
      <div style="padding:0 14px 10px 14px;">
        <div id="hud-status" style="font-size:12px;color:#444;margin-bottom:8px;">Starting…</div>
        <div style="height:10px;background:#f3f4f6;border-radius:8px;overflow:hidden;">
          <div id="hud-bar" style="height:100%;width:5%;background:linear-gradient(90deg,#2563eb,#3b82f6);transition:width .35s ease;"></div>
        </div>
      </div>
      <div id="hud-settings" style="display:none;padding:10px 14px;border-top:1px solid #f1f5f9;background:#fafafa;">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px 16px;">
          <label style="display:flex;align-items:center;gap:6px;"><input type="checkbox" id="set-autoDisable"> Auto-disable on 0 flows (batch only)</label>
          <label style="display:flex;align-items:center;gap:6px;"><input type="checkbox" id="set-reloadAfter"> Reload after disable</label>
          <label style="display:flex;align-items:center;gap:6px;"><input type="checkbox" id="set-tighten"> Enable tightening (flows > 0)</label>
          <span></span>
          <label style="display:flex;align-items:center;gap:6px;"><input type="checkbox" id="set-verbose"> Verbose console logs</label>
          <label style="display:flex;align-items:center;gap:6px;"><input type="checkbox" id="set-verbosePoll"> Log each poll (throttled)</label>
          <label style="display:flex;align-items:center;gap:6px;"><input type="checkbox" id="set-logRaw"> Raw JSON strings</label>
          <label style="display:flex;align-items:center;gap:6px;"><input type="checkbox" id="set-toasts"> Show toasts</label>
          <label style="display:flex;align-items:center;gap:6px;"><input type="checkbox" id="set-hudEnabled"> HUD enabled</label>
          <span></span>
          <label>Lookback days <input type="number" id="set-lookback" min="1" max="365" step="1" style="width:90px;margin-left:6px;"></label>
          <label>Poll interval (ms) <input type="number" id="set-poll" min="500" max="30000" step="100" style="width:110px;margin-left:6px;"></label>
          <label>Max poll time (ms) <input type="number" id="set-maxpoll" min="10000" max="900000" step="500" style="width:120px;margin-left:6px;"></label>
          <label>Max results <input type="number" id="set-maxresults" min="1000" max="500000" step="1000" style="width:120px;margin-left:6px;"></label>
        </div>
        <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:10px;">
          <button id="hud-settings-reset" style="border:1px solid #d1d5db;background:#fff;border-radius:8px;padding:6px 12px;cursor:pointer;">Reset</button>
          <button id="hud-settings-save"  style="border:1px solid #10b981;background:#10b981;color:#fff;border-radius:8px;padding:6px 12px;cursor:pointer;">Save</button>
        </div>
      </div>
      <div id="hud-decision" style="display:none;padding:10px 14px;border-top:1px solid #f1f5f9;display:flex;gap:8px;justify-content:flex-end;">
        <button id="hud-disable" style="border:1px solid #0ea5e9;background:#0284c7;color:#fff;border-radius:8px;padding:6px 12px;cursor:pointer;">${disableBtnText}</button>
      </div>
      <div id="hud-tighten" style="display:none;padding:10px 14px;border-top:1px solid #f1f5f9;">
        <div style="font-weight:600;margin-bottom:6px;">Flows-based tightening proposal</div>
        <div id="hud-tighten-body" style="font-size:12px;color:#374151;"></div>
        <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:10px;">
          <button id="hud-tighten-apply" style="border:1px solid #10b981;background:#10b981;color:#fff;border-radius:8px;padding:6px 12px;cursor:pointer;">Apply</button>
          <button id="hud-tighten-skip"  style="border:1px solid #d1d5db;background:#fff;border-radius:8px;padding:6px 12px;cursor:pointer;">Skip</button>
        </div>
      </div>
    `;
    if (!document.body) { window.addEventListener('DOMContentLoaded', () => document.body.appendChild(el)); }
    else document.body.appendChild(el);

    HUD = el;
    HUD_STATUS = el.querySelector('#hud-status');
    HUD_RULEID  = el.querySelector('#hud-ruleid');
    HUD_FLOWS   = el.querySelector('#hud-flows');
    HUD_BAR     = el.querySelector('#hud-bar');
    HUD_DECISION = el.querySelector('#hud-decision');
    HUD_BTN_DISABLE = el.querySelector('#hud-disable');
    HUD_SETTINGS_WRAP = el.querySelector('#hud-settings');
    HUD_GEAR_BTN = el.querySelector('#hud-gear');
    HUD_TIGHTEN = el.querySelector('#hud-tighten');
    HUD_TIGHTEN_APPLY = el.querySelector('#hud-tighten-apply');
    HUD_TIGHTEN_SKIP  = el.querySelector('#hud-tighten-skip');

    el.querySelector('#hud-close').onclick = () => removeHUD();
    el.querySelector('#hud-cancel').onclick = () => { ABORT_REQUESTED = true; setHUDStep('cancelled'); };

    hydrateSettingsForm();
    wireSettingsForm();
    setDisableButtonEnabled(false);
  }
  function removeHUD(){
    HUD?.remove();
    HUD=HUD_STATUS=HUD_RULEID=HUD_FLOWS=HUD_BAR=null;
    HUD_DECISION=HUD_BTN_DISABLE=HUD_SETTINGS_WRAP=HUD_GEAR_BTN=null;
    HUD_TIGHTEN=HUD_TIGHTEN_APPLY=HUD_TIGHTEN_SKIP=null;
  }
  function setHUDRuleId(id){ if (HUD_RULEID) HUD_RULEID.textContent=id||'—'; }
  function setHUDStatus(k){ if (!HUD_STATUS) return; const s=STEPS[k]; HUD_STATUS.textContent=s?s.label:(k||''); }
  function setHUDStep(k){ if(!HUD_STATUS||!HUD_BAR) return; const s=STEPS[k]||{pct:100,label:k}; HUD_STATUS.textContent=s.label; HUD_BAR.style.width=`${s.pct}%`; }
  function nudgePollProgress(t){ if (!HUD_BAR) return; const base=70, span=20; HUD_BAR.style.width=`${base+(t%10)*(span/10)}%`; }
  function setHUDFlows(n){ if (HUD_FLOWS) HUD_FLOWS.textContent=(n??'—'); }
  function showDecisionRow(show){ if (HUD_DECISION) HUD_DECISION.style.display=show?'flex':'none'; }
  function setDisableButtonEnabled(en){ if(!HUD_BTN_DISABLE) return; HUD_BTN_DISABLE.disabled=!en; HUD_BTN_DISABLE.style.opacity=en?'1':'0.5'; HUD_BTN_DISABLE.style.cursor=en?'pointer':'not-allowed'; }
  function updateDisableButtonText(){ if (HUD_BTN_DISABLE) HUD_BTN_DISABLE.textContent = SETTINGS.reloadAfterDisable ? 'Disable & Reload' : 'Disable'; }
  function showTightenPanel(html, show){
    if (!HUD_TIGHTEN) return;
    HUD_TIGHTEN.style.display = show ? 'block' : 'none';
    const body = HUD_TIGHTEN.querySelector('#hud-tighten-body');
    if (body && typeof html === 'string') body.innerHTML = html;
  }

  /** Ruleset Batch HUD */
  function createBatchHUD(total){
    removeBatchHUD();
    const el = document.createElement('div');
    el.id='illumio-rulecore-batch-hud';
    el.style.cssText = `position:fixed;left:16px;bottom:16px;width:620px;background:#fff;color:#111;z-index:100000;border:1px solid #e5e7eb;border-radius:12px;box-shadow:0 14px 30px rgba(0,0,0,0.18);font:13px/1.4 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;overflow:hidden;`;
    el.innerHTML = `
      <div style="padding:12px 14px;border-bottom:1px solid #eee;display:flex;align-items:center;justify-content:space-between;">
        <div style="display:flex;align-items:center;gap:10px;">
          <div style="font-weight:700;">Ruleset Review — Batch</div>
        </div>
        <div style="display:flex;gap:8px;">
          <button id="bhud-cancel" style="border:1px solid #ef4444;background:#fff;color:#ef4444;border-radius:8px;padding:4px 10px;cursor:pointer;">Cancel</button>
          <button id="bhud-close"  style="border:1px solid #d1d5db;background:#fff;border-radius:8px;padding:4px 10px;cursor:pointer;">Close</button>
        </div>
      </div>
      <div style="padding:10px 14px;">
        <div id="bhud-status" style="font-size:12px;color:#444;margin-bottom:6px;">Preparing…</div>
        <div id="bhud-counters" style="font-size:12px;color:#374151;margin-bottom:6px;">0 / ${total} processed · 0 skipped · 0 failed · 0 disabled · 0 tightened</div>
        <div id="bhud-current" style="font-size:12px;color:#6b7280;margin-bottom:8px;">Current: —</div>
        <div style="height:10px;background:#f3f4f6;border-radius:8px;overflow:hidden;">
          <div id="bhud-bar" style="height:100%;width:0%;background:linear-gradient(90deg,#059669,#10b981);transition:width .35s ease;"></div>
        </div>
        <details id="bhud-details" style="margin-top:12px;">
          <summary style="cursor:pointer;font-weight:600;">Batch summary details</summary>
          <div style="margin-top:8px;display:flex;flex-direction:column;gap:6px;"></div>
        </details>
      </div>
    `;
    if (!document.body) { window.addEventListener('DOMContentLoaded', () => document.body.appendChild(el)); }
    else document.body.appendChild(el);
    BHUD = el;
    BHUD_STATUS   = el.querySelector('#bhud-status');
    BHUD_COUNTERS = el.querySelector('#bhud-counters');
    BHUD_CURR     = el.querySelector('#bhud-current');
    BHUD_BAR      = el.querySelector('#bhud-bar');
    BHUD_DETAILS  = el.querySelector('#bhud-details > div');

    el.querySelector('#bhud-close').onclick = () => removeBatchHUD();
    el.querySelector('#bhud-cancel').onclick = () => { BATCH_ABORT_REQUESTED = true; if (BHUD_STATUS) BHUD_STATUS.textContent = 'Cancelling after current rule…'; };
  }
  function removeBatchHUD(){ BHUD?.remove(); BHUD=BHUD_STATUS=BHUD_COUNTERS=BHUD_BAR=BHUD_CURR=BHUD_DETAILS=null; }
  function updateBatchHUDStatus(text){ if (BHUD_STATUS) BHUD_STATUS.textContent = String(text||''); }
  function updateBatchHUDCurrent(text){ if (BHUD_CURR) BHUD_CURR.textContent = `Current: ${text||'—'}`; }
  function updateBatchHUDCounters({processed, total, skipped, failed, disabled, tightened}){
    if (BHUD_COUNTERS) BHUD_COUNTERS.textContent = `${processed} / ${total} processed · ${skipped} skipped · ${failed} failed · ${disabled} disabled · ${tightened} tightened`;
    if (BHUD_BAR) BHUD_BAR.style.width = `${Math.round((processed/Math.max(1,total))*100)}%`;
  }
  function appendBatchDetail(html){
    if (!BHUD_DETAILS) return;
    const div = document.createElement('div');
    div.innerHTML = html;
    div.style.cssText = 'padding:8px;border:1px solid #eef2f7;border-radius:6px;background:#fafafa;font-size:12px;';
    BHUD_DETAILS.appendChild(div);
  }

  /** Policies Batch HUD */
  function createPoliciesHUD(totalRulesets){
    removePoliciesHUD();
    const el = document.createElement('div');
    el.id='illumio-policies-batch-hud';
    el.style.cssText = `position:fixed;left:16px;bottom:16px;width:680px;background:#fff;color:#111;z-index:100000;border:1px solid #e5e7eb;border-radius:12px;box-shadow:0 14px 30px rgba(0,0,0,0.18);font:13px/1.4 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;overflow:hidden;`;
    el.innerHTML = `
      <div style="padding:12px 14px;border-bottom:1px solid #eee;display:flex;align-items:center;justify-content:space-between;">
        <div style="display:flex;align-items:center;gap:10px;">
          <div style="font-weight:700;">Policies — Selected Rulesets Review</div>
        </div>
        <div style="display:flex;gap:8px;">
          <button id="phud-cancel" style="border:1px solid #ef4444;background:#fff;color:#ef4444;border-radius:8px;padding:4px 10px;cursor:pointer;">Cancel</button>
          <button id="phud-close"  style="border:1px solid #d1d5db;background:#fff;border-radius:8px;padding:4px 10px;cursor:pointer;">Close</button>
        </div>
      </div>
      <div style="padding:10px 14px;">
        <div id="phud-status" style="font-size:12px;color:#444;margin-bottom:6px;">Preparing…</div>
        <div style="display:flex;gap:18px;flex-wrap:wrap;margin-bottom:6px;">
          <div id="phud-counters-rs" style="font-size:12px;color:#374151;">Rulesets: 0 / ${totalRulesets} processed · 0 skipped · 0 failed</div>
          <div id="phud-counters-rules" style="font-size:12px;color:#374151;">Rules: 0 processed · 0 skipped · 0 failed · 0 disabled · 0 tightened</div>
        </div>
        <div id="phud-current-rs" style="font-size:12px;color:#6b7280;margin-bottom:2px;">Current ruleset: —</div>
        <div id="phud-current-rule" style="font-size:12px;color:#6b7280;margin-bottom:8px;">Current rule: —</div>
        <div style="height:10px;background:#f3f4f6;border-radius:8px;overflow:hidden;">
          <div id="phud-bar" style="height:100%;width:0%;background:linear-gradient(90deg,#3b82f6,#60a5fa);transition:width .35s ease;"></div>
        </div>
        <details id="phud-details" style="margin-top:12px;">
          <summary style="cursor:pointer;font-weight:600;">Run details</summary>
          <div style="margin-top:8px;display:flex;flex-direction:column;gap:6px;"></div>
        </details>
      </div>
    `;
    document.body.appendChild(el);
    PHUD = el;
    PHUD_STATUS   = el.querySelector('#phud-status');
    PHUD_BAR      = el.querySelector('#phud-bar');
    PHUD_COUNTERS_RS = el.querySelector('#phud-counters-rs');
    PHUD_COUNTERS_RULES = el.querySelector('#phud-counters-rules');
    PHUD_CURR_RS  = el.querySelector('#phud-current-rs');
    PHUD_CURR_RULE= el.querySelector('#phud-current-rule');
    PHUD_DETAILS  = el.querySelector('#phud-details > div');

    el.querySelector('#phud-close').onclick = () => removePoliciesHUD();
    el.querySelector('#phud-cancel').onclick = () => { POLICIES_ABORT_REQUESTED = true; if (PHUD_STATUS) PHUD_STATUS.textContent = 'Cancelling after current rule…'; };
  }
  function removePoliciesHUD(){ PHUD?.remove(); PHUD=PHUD_STATUS=PHUD_BAR=PHUD_COUNTERS_RS=PHUD_COUNTERS_RULES=PHUD_CURR_RS=PHUD_CURR_RULE=PHUD_DETAILS=null; }
  function updatePoliciesHUDStatus(text){ if (PHUD_STATUS) PHUD_STATUS.textContent = String(text||''); }
  function updatePoliciesHUDCurrentRS(text){ if (PHUD_CURR_RS) PHUD_CURR_RS.textContent = `Current ruleset: ${text||'—'}`; }
  function updatePoliciesHUDCurrentRule(text){ if (PHUD_CURR_RULE) PHUD_CURR_RULE.textContent = `Current rule: ${text||'—'}`; }
  function updatePoliciesHUDCounters({rsProcessed, rsTotal, rsSkipped, rsFailed, rulesProcessed, rulesSkipped, rulesFailed, rulesDisabled, rulesTightened}) {
    if (PHUD_COUNTERS_RS) PHUD_COUNTERS_RS.textContent = `Rulesets: ${rsProcessed} / ${rsTotal} processed · ${rsSkipped} skipped · ${rsFailed} failed`;
    if (PHUD_COUNTERS_RULES) PHUD_COUNTERS_RULES.textContent = `Rules: ${rulesProcessed} processed · ${rulesSkipped} skipped · ${rulesFailed} failed · ${rulesDisabled} disabled · ${rulesTightened} tightened`;
    const pct = rsTotal ? Math.round((rsProcessed/rsTotal)*100) : 0;
    if (PHUD_BAR) PHUD_BAR.style.width = `${pct}%`;
  }
  function appendPoliciesDetail(html){
    if (!PHUD_DETAILS) return;
    const div = document.createElement('div');
    div.innerHTML = html;
    div.style.cssText = 'padding:8px;border:1px solid #eef2f7;border-radius:6px;background:#fafafa;font-size:12px;';
    PHUD_DETAILS.appendChild(div);
  }

  /***************************************************************************
   * SETTINGS PANEL PLUMBING
   ***************************************************************************/
  function hydrateSettingsForm(){ if(!HUD_SETTINGS_WRAP) return; const $=(s)=>HUD_SETTINGS_WRAP.querySelector(s);
    $('#set-autoDisable').checked=!!SETTINGS.autoDisableOnZeroFlows;
    $('#set-reloadAfter').checked=!!SETTINGS.reloadAfterDisable;
    $('#set-tighten').checked=!!SETTINGS.tightenIpListsEnabled;
    $('#set-verbose').checked=!!SETTINGS.verbose;
    $('#set-verbosePoll').checked=!!SETTINGS.verbosePoll;
    $('#set-logRaw').checked=!!SETTINGS.logRawJson;
    $('#set-toasts').checked=!!SETTINGS.showToasts;
    $('#set-hudEnabled').checked=!!SETTINGS.hudEnabled;
    $('#set-lookback').value=String(SETTINGS.lookbackDays);
    $('#set-poll').value=String(SETTINGS.pollIntervalMs);
    $('#set-maxpoll').value=String(SETTINGS.maxPollMs);
    $('#set-maxresults').value=String(SETTINGS.maxResults);
  }
  function wireSettingsForm(){ if(!HUD_SETTINGS_WRAP) return; const $=(s)=>HUD_SETTINGS_WRAP.querySelector(s); const save=$('#hud-settings-save'), reset=$('#hud-settings-reset');
    reset.onclick=()=>{ Object.assign(SETTINGS,{...DEFAULTS}); saveSettings(); hydrateSettingsForm(); updateDisableButtonText(); toast('Settings reset to defaults.','ok'); if(!SETTINGS.hudEnabled) removeHUD(); };
    save.onclick=()=>{
      SETTINGS.autoDisableOnZeroFlows=$('#set-autoDisable').checked;
      SETTINGS.reloadAfterDisable=$('#set-reloadAfter').checked;
      SETTINGS.tightenIpListsEnabled=$('#set-tighten').checked;
      SETTINGS.verbose=$('#set-verbose').checked;
      SETTINGS.verbosePoll=$('#set-verbosePoll').checked;
      SETTINGS.logRawJson=$('#set-logRaw').checked;
      SETTINGS.showToasts=$('#set-toasts').checked;
      SETTINGS.hudEnabled=$('#set-hudEnabled').checked;
      const clamp=(v,min,max,fb)=>{ const n=parseInt(v,10); return Number.isFinite(n)?Math.min(max,Math.max(min,n)):fb; };
      SETTINGS.lookbackDays=clamp($('#set-lookback').value,1,365,DEFAULTS.lookbackDays);
      SETTINGS.pollIntervalMs=clamp($('#set-poll').value,500,30000,DEFAULTS.pollIntervalMs);
      SETTINGS.maxPollMs=clamp($('#set-maxpoll').value,10000,900000,DEFAULTS.maxPollMs);
      SETTINGS.maxResults=clamp($('#set-maxresults').value,1000,500000,DEFAULTS.maxResults);
      saveSettings(); updateDisableButtonText(); toast('Settings saved.','ok'); if(!SETTINGS.hudEnabled) setTimeout(()=>removeHUD(),300);
    };
  }
  function escapeHtml(s){
    return String(s ?? '')
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;')
      .replace(/'/g,'&#039;');
  }

  /***************************************************************************
   * CONTEXT / ORG / CSRF
   ***************************************************************************/
  function isRulesetDetailPage(){ return /#\/rulesets\/\d+/.test(location.hash); }
  function getRulesetIdFromUrl(){ const m=location.hash.match(/\/rulesets\/(\d+)/); return m?m[1]:null; }

  let ORG_ID=null;
  const baseUrl=location.origin;

  function detectOrgIdFromUrlOrState() {
    if (ORG_ID) return ORG_ID;
    try {
      const m = String(location.pathname).match(/\/orgs\/(\d+)/);
      if (m) ORG_ID = m[1];
      else if (window.__INITIAL_STATE__?.organization?.id) ORG_ID = String(window.__INITIAL_STATE__.organization.id);
      else {
        const meta = document.querySelector('meta[name="org-id"]')?.content;
        if (meta) ORG_ID = String(meta);
      }
    } catch {}
    return ORG_ID;
  }

  (function wrapTransport() {
    try {
      const origFetch = window.fetch?.bind(window);
      if (origFetch) {
        window.fetch = async (...args) => {
          const url = typeof args[0] === 'string' ? args[0] : args[0]?.url;
          if (!ORG_ID && url) {
            const m = String(url).match(/\/orgs\/(\d+)\//);
            if (m) ORG_ID = m[1];
          }
          return origFetch(...args);
        };
      }
      if (window.XMLHttpRequest) {
        const origOpen = XMLHttpRequest.prototype.open;
        XMLHttpRequest.prototype.open = new Proxy(origOpen, {
          apply(target, thisArg, args) {
            try {
              const url = args?.[1];
              if (!ORG_ID && url) {
                const m = String(url).match(/\/orgs\/(\d+)\//);
                if (m) ORG_ID = m[1];
              }
            } catch {}
            return Reflect.apply(target, thisArg, args);
          }
        });
      }
    } catch {}
  })();

  function getCsrfToken(){
    try {
      const meta=document.querySelector('meta[name="csrf-token"]')?.content; if (meta) return meta;
      const ck=document.cookie.match(/csrf_token=([^;]+)/); if (ck) return ck[1];
      const alt=document.cookie.match(/CSRF-TOKEN=([^;]+)/); if (alt) return decodeURIComponent(alt[1]);
      if (window.__INITIAL_STATE__?.csrfToken) return window.__INITIAL_STATE__.csrfToken;
      if (window._csrfToken) return window._csrfToken;
    } catch {}
    return null;
  }

  /***************************************************************************
   * API
   ***************************************************************************/
  async function apiGet(path){
    const url = path.startsWith('/api/v2') ? `${baseUrl}${path}` : `${baseUrl}/api/v2${path}`;
    const res=await fetch(url,{credentials:'include'}); const text=await res.text();
    if (!res.ok) throw new Error(`GET ${path} HTTP ${res.status}: ${text}`);
    logRaw(`GET ${path} RESPONSE`, text);
    return text?JSON.parse(text):null;
  }
  async function apiPut(path,payload){
    const csrf=getCsrfToken(); if(!csrf) throw new Error('CSRF token not found');
    const url=`${baseUrl}/api/v2${path}`; const body=JSON.stringify(payload);
    logRaw(`PUT ${path} PAYLOAD`, body);
    const res=await fetch(url,{method:'PUT',credentials:'include',headers:{'Content-Type':'application/json','Accept':'application/json','x-csrf-token':csrf},body});
    const text=await res.text(); logRaw(`PUT ${path} RESPONSE`, text);
    if (!res.ok) throw new Error(`PUT ${path} HTTP ${res.status}: ${text}`);
    return text?JSON.parse(text):null;
  }
  async function fetchRuleset(rsId){ return apiGet(`/orgs/${ORG_ID}/sec_policy/draft/rule_sets/${rsId}`); }
  async function listRulesets(max = 5000){
    return apiGet(`/orgs/${ORG_ID}/sec_policy/draft/rule_sets?max_results=${max}`);
  }

  /***************************************************************************
   * DOM → API signature mapping
   ***************************************************************************/
  function uiHrefToApi(href){
    if (!href || !ORG_ID) return null;
    if (href.endsWith('#/workloads')) return `/orgs/${ORG_ID}/workloads`;
    const m=href.match(/#\/([^/]+)\/([^/?]+)/); if (!m) return null;
    const [,type,id]=m;
    const map={ workloads:'workloads', labels:'labels', iplists:'sec_policy/draft/ip_lists', services:'sec_policy/draft/services' };
    return map[type]?`/orgs/${ORG_ID}/${map[type]}/${id}`:null;
  }
  function extractColumnHrefs(row,tid){
    const anchors=[...row.querySelectorAll(`[data-tid="${tid}"] a[data-tid^="comp-pill"]`)];
    const hrefs=anchors.map(a=>uiHrefToApi(a.href)).filter(Boolean); hrefs.sort(); return hrefs;
  }
  function domSignature(row){
    const consumers=extractColumnHrefs(row,'comp-grid-column-consumers');
    const providers=extractColumnHrefs(row,'comp-grid-column-providers');
    const services = extractColumnHrefs(row,'comp-grid-column-providingservices');
    const extraTxt = row.querySelector('[data-tid="comp-grid-column-extrascope"]')?.innerText || '';
    const unscoped = /extra/i.test(extraTxt);
    const sig={consumers,providers,services,unscoped_consumers:unscoped}; logRaw('DOM Signature', sig); return sig;
  }
  function rawApiSignature(rule){
    const side=(arr)=>(arr||[]).map(e =>
      e.workload?.href || e.ip_list?.href || e.label?.href || (e.actors==='ams' ? `/orgs/${ORG_ID}/workloads` : null)
    ).filter(Boolean).sort();
    const ing=(rule.ingress_services||[]).map(s=>s.href).filter(Boolean);
    const eg =(rule.egress_services ||[]).map(s=>s.href).filter(Boolean);
    const sig={consumers:side(rule.consumers),providers:side(rule.providers),services:[...new Set([...ing,...eg])].sort(),unscoped_consumers:!!rule.unscoped_consumers};
    logRaw('Rule RAW Signature', sig); return sig;
  }
  function getRulesetScopeClauses(ruleset){
    const out=[]; for (const clause of ruleset?.scopes||[]) {
      if (!Array.isArray(clause) || !clause.length) continue;
      const hrefs=clause.map(s=>s?.label?.href).filter(Boolean); if (hrefs.length) out.push(hrefs);
    } logRaw('Ruleset Scope Clauses', out); return out;
  }
  function getSingleScopeLabels(ruleset){ const first=getRulesetScopeClauses(ruleset)[0]||[]; logRaw('Using Scope Labels (single-clause)', first); return first; }
  function isAllWorkloads(side){ return side.length===1 && side[0].endsWith('/workloads'); }
  function applyScopeToSide(side,scopeLabels){
    if (!scopeLabels?.length) return side;
    if (isAllWorkloads(side)) return [...scopeLabels].sort();
    if (side.every(h=>h.includes('/labels/'))) return [...new Set([...side,...scopeLabels])].sort();
    return side;
  }
  function effectiveApiSignature(raw,rule,scopeLabels){
    const eff={consumers:[...raw.consumers],providers:[...raw.providers],services:[...raw.services],unscoped_consumers:raw.unscoped_consumers};
    if (rule.unscoped_consumers === false) {
      eff.consumers = applyScopeToSide(eff.consumers, scopeLabels);
      eff.providers = applyScopeToSide(eff.providers, scopeLabels);
    }
    if (rule.unscoped_consumers === true) {
      eff.providers = applyScopeToSide(eff.providers, scopeLabels);
    }
    eff.consumers.sort(); eff.providers.sort(); eff.services.sort();
    logRaw('Rule EFFECTIVE Signature', eff); return eff;
  }
  function matchDomToRule(domSig,rule){
    const raw=rawApiSignature(rule);
    const sidesMatch =
      JSON.stringify(domSig.consumers)===JSON.stringify(raw.consumers) &&
      JSON.stringify(domSig.providers)===JSON.stringify(raw.providers) &&
      domSig.unscoped_consumers===raw.unscoped_consumers;
    if (!sidesMatch) return false;
    if (domSig.services?.length) return JSON.stringify(domSig.services)===JSON.stringify(raw.services);
    return true;
  }

  /***************************************************************************
   * Explorer async + services building (incl. ranges)
   ***************************************************************************/
  function hrefToTrafficEntity(href){
    if (!href) return null;
    if (href.includes('/ip_lists/')) return {ip_list:{href}};
    if (href.includes('/labels/'))   return {label:{href}};
    if (href.endsWith('/workloads')) return {actors:'ams'};
    if (href.includes('/workloads/'))return {workload:{href}};
    return null;
  }
  function sideToExplorerInclude(sideHrefs){
    if ((sideHrefs||[]).some(h=>h.endsWith('/workloads'))) return [];
    const entities=(sideHrefs||[]).map(hrefToTrafficEntity).filter(e=>e && !e.actors);
    return entities.length ? [entities] : [];
  }

  function pickNumber(...vals) { for (const v of vals) if (Number.isFinite(v)) return v; return null; }
  function normalizeInlineService(item) {
    const proto = pickNumber(item?.proto, item?.protocol);
    const port  = pickNumber(item?.port, item?.dst_port, item?.dst_port_range_start, item?.from_port, item?.port_range_start);
    const to    = pickNumber(item?.to_port, item?.dst_port_range_end, item?.port_range_end);
    if (!Number.isFinite(proto) || !Number.isFinite(port)) return null;
    const out = { proto, port };
    if (Number.isFinite(to) && to >= port) out.to_port = to;
    return out;
  }
  async function fetchServiceAndExpandInline(href) {
    try {
      const svc = await apiGet(href);
      const ports = Array.isArray(svc?.service_ports) ? svc.service_ports : (Array.isArray(svc?.ports) ? svc.ports : []);
      const out = [];
      for (const p of ports) {
        const proto = pickNumber(p?.proto, p?.protocol);
        const start = pickNumber(p?.port, p?.from_port, p?.port_range_start);
        const end   = pickNumber(p?.to_port, p?.port_range_end);
        if (!Number.isFinite(proto) || !Number.isFinite(start)) continue;
        if (Number.isFinite(end) && end >= start) out.push({ proto, port: start, to_port: end });
        else out.push({ proto, port: start });
      }
      return out;
    } catch (e) { warn('Service expand failed for href:', href, e); return []; }
  }
  async function buildServicesIncludeFromRuleAsync(rule) {
    const raw = [
      ...(Array.isArray(rule?.ingress_services) ? rule.ingress_services : []),
      ...(Array.isArray(rule?.egress_services)  ? rule.egress_services  : [])
    ];
    if (!raw.length) return [];
    const seen = new Set(); const out = [];

    // Inline
    for (const s of raw) {
      const inl = normalizeInlineService(s); if (inl) {
        const key = `${inl.proto}:${inl.port}-${inl.to_port ?? inl.port}`;
        if (!seen.has(key)) { seen.add(key); out.push(inl); }
      }
    }
    // Hrefs
    for (const s of raw) {
      const href = s?.href || s?.service?.href; if (!href) continue;
      const expanded = await fetchServiceAndExpandInline(href);
      for (const it of expanded) {
        const key = `${it.proto}:${it.port}-${it.to_port ?? it.port}`;
        if (!seen.has(key)) { seen.add(key); out.push(it); }
      }
    }
    return out;
  }

  async function runTrafficQuery(effSig, ruleName, currentMatchedRule){
    const csrf = getCsrfToken(); if(!csrf) throw new Error('CSRF token not found');
    const servicesInclude = await buildServicesIncludeFromRuleAsync(currentMatchedRule);
    logRaw('Services include (derived from rule)', servicesInclude);

    const payload={
      sources:{ include: sideToExplorerInclude(effSig.consumers), exclude:[] },
      destinations:{ include: sideToExplorerInclude(effSig.providers), exclude:[] },
      services:{ include: servicesInclude, exclude:[] },
      sources_destinations_query_op:'and',
      start_date:new Date(Date.now()-SETTINGS.lookbackDays*24*60*60*1000).toISOString(),
      end_date:new Date().toISOString(),
      policy_decisions:[], boundary_decisions:[],
      query_name:`RULE_CORE_TRAFFIC_${ruleName||'Unnamed'}`,
      exclude_workloads_from_ip_list_query:false,
      max_results:SETTINGS.maxResults
    };
    const body=JSON.stringify(payload,null,2); logRaw('AsyncQuery CREATE payload',body);

    const createRes=await fetch(`${baseUrl}/api/v2/orgs/${ORG_ID}/traffic_flows/async_queries`,{
      method:'POST',credentials:'include',headers:{'Content-Type':'application/json','x-csrf-token':csrf},body
    });
    const createText=await createRes.text(); logRaw('AsyncQuery CREATE response',createText);
    if(!createRes.ok) throw new Error(`Async create HTTP ${createRes.status}: ${createText}`);

    let created; try{created=createText?JSON.parse(createText):null;}catch{created=null;}
    const queryHref=created?.href; if(!queryHref) throw new Error('Async query did not return href');

    const start=Date.now(); let polls=0; let lastStatus='';
    while(true){
      if(BATCH_MODE ? BATCH_ABORT_REQUESTED : ABORT_REQUESTED) return {status:'aborted',flowsCount:null,queryHref,downloadPath:null};
      if(Date.now()-start>SETTINGS.maxPollMs) return {status:'timeout',flowsCount:null,queryHref,downloadPath:null};
      const res=await fetch(`${baseUrl}/api/v2${queryHref}`,{credentials:'include'}); const text=await res.text();
      if(SETTINGS.verbosePoll){ try{ const json=text?JSON.parse(text):null; const s=json?.status||''; if(s!==lastStatus){logRaw(`AsyncQuery POLL #${polls} (status change)`,text); lastStatus=s;} else if(polls%3===0){logRaw(`AsyncQuery POLL #${polls}`,text);} } catch{ logRaw(`AsyncQuery POLL #${polls}`,text); } }
      if(!res.ok) return {status:'http_error',flowsCount:null,queryHref,downloadPath:null};
      let data; try{ data=text?JSON.parse(text):null; } catch{ data=null; }
      const status=data?.status; polls++;
      if(!status || ['queued','working','pending','running'].includes(status)){ await new Promise(r=>setTimeout(r,SETTINGS.pollIntervalMs)); continue; }
      return {status:'completed',flowsCount:data?.flows_count??0,queryHref,downloadPath:data?.result||null};
    }
  }

  async function downloadAsyncQueryResultsJson(queryHref, downloadPath){
    let url=null;
    if (downloadPath) url=`${baseUrl}/api/v2${String(downloadPath).replace(/^\/?api\/v2/,'')}`;
    else {
      const uuid=queryHref?.split('/').pop();
      if(!uuid) return [];
      url=`${baseUrl}/api/v2/orgs/${ORG_ID}/traffic_flows/async_queries/${uuid}/download`;
    }
    const res=await fetch(url,{credentials:'include',headers:{'Accept':'application/json'}});
    const text=await res.text(); logRaw('Flows DOWNLOAD response',text);
    if(!res.ok) throw new Error(`Download failed HTTP ${res.status}: ${text}`);
    const json=text?JSON.parse(text):[];
    return Array.isArray(json)?json:(Array.isArray(json?.items)?json.items:[]);
  }

  async function disableRule(ruleHref){
    const current=await apiGet(ruleHref);
    const payload={
      providers:current.providers??[],
      consumers:current.consumers??[],
      enabled:false,
      ingress_services:current.ingress_services??[],
      egress_services:current.egress_services??[],
      network_type:current.network_type??'',
      description:current.description??'',
      consuming_security_principals:current.consuming_security_principals??[],
      sec_connect:current.sec_connect??false,
      machine_auth:current.machine_auth??false,
      stateless:current.stateless??false,
      unscoped_consumers:current.unscoped_consumers??false,
      use_workload_subnets:current.use_workload_subnets??[],
      resolve_labels_as:current.resolve_labels_as??{}
    };
    await apiPut(ruleHref,payload);
    return true;
  }

  /***************************************************************************
   * Tightening helpers
   ***************************************************************************/
  function buildIndexStats(flows, side){
    const stats=new Map();
    for (const f of flows||[]) {
      const arr=f?.[side]?.ip_lists; if(!Array.isArray(arr)) continue;
      for (let i=0;i<arr.length;i++){
        const it=arr[i]; const href=it?.href; if(!href) continue;
        if(!stats.has(href)) stats.set(href,{name:it?.name||'', indexSum:0, indexCount:0, maxIndex:-Infinity});
        const s=stats.get(href); s.indexSum+=i; s.indexCount+=1; if(i>s.maxIndex) s.maxIndex=i;
      }
    }
    return stats;
  }
  function computeCommonHrefSet(flows, side){
    let common=null, considered=0;
    for (const f of flows||[]) {
      const arr=f?.[side]?.ip_lists; if(!Array.isArray(arr) || !arr.length) continue;
      considered++;
      const set=new Set(arr.map(x=>x?.href).filter(Boolean));
      if (common==null) common=set; else { const next=new Set(); for(const h of common) if (set.has(h)) next.add(h); common=next; }
      if (common.size===0) break;
    }
    return { common: common||new Set(), consideredFlows: considered };
  }
  function chooseMinimaxFromSet(commonSet, indexStats){
    if (!commonSet || commonSet.size===0) return null;
    let bestHref=null, bestMax=Number.POSITIVE_INFINITY, bestAvg=Number.POSITIVE_INFINITY;
    for (const href of commonSet){
      const s=indexStats.get(href); if(!s || s.indexCount===0) continue;
      const maxIdx=s.maxIndex; const avgIdx=s.indexSum/Math.max(1,s.indexCount);
      if (maxIdx<bestMax) { bestMax=maxIdx; bestAvg=avgIdx; bestHref=href; }
      else if (maxIdx===bestMax) {
        if (avgIdx<bestAvg) { bestAvg=avgIdx; bestHref=href; }
        else if (avgIdx===bestAvg && bestHref && href<bestHref) bestHref=href;
      }
    }
    return bestHref;
  }
  function isIPv6(ip) { return typeof ip === 'string' && ip.includes(':'); }
  function partitionFlowsByFamily(flows) {
    const v4=[], v6=[];
    for (const f of flows||[]) {
      const ip = f?.src?.ip;
      if (!ip) continue;
      if (isIPv6(ip)) v6.push(f); else v4.push(f);
    }
    return { v4Flows:v4, v6Flows:v6 };
  }

  function extractIpListHrefsFromRule(rule){
    const side = (arr) => (arr || []).map(a => a?.ip_list?.href).filter(Boolean);
    return {
      consumerIpLists: [...new Set(side(rule.consumers))],
      providerIpLists: [...new Set(side(rule.providers))]
    };
  }

  const ipListNameCache = new Map(); // href -> friendly name
  async function ensureIpListNames(hrefs) {
    const tasks = [];
    for (const href of (hrefs || [])) {
      if (!href || ipListNameCache.has(href)) continue;
      tasks.push((async () => {
        try {
          const obj = await apiGet(href);
          const name = obj?.name || href;
          ipListNameCache.set(href, name);
        } catch {
          ipListNameCache.set(href, href);
        }
      })());
    }
    if (tasks.length) await Promise.all(tasks);
  }

  function renderTightenProposalHtml(proposals){
    const section=(title,p)=>{
      const map = new Map(p.nameMap ? p.nameMap : []);
      if (p.oldHrefs) {
        for (const h of p.oldHrefs) {
          const cached = ipListNameCache.get(h);
          if (cached) {
            const prev = map.get(h) || { name: '', indexSum: 0, indexCount: 0, maxIndex: -Infinity };
            prev.name = prev.name || cached;
            map.set(h, prev);
          }
        }
      }
      if (p.newHref) {
        const cached = ipListNameCache.get(p.newHref);
        if (cached) {
          const prev = map.get(p.newHref) || { name: '', indexSum: 0, indexCount: 0, maxIndex: -Infinity };
          prev.name = prev.name || cached;
          map.set(p.newHref, prev);
        }
      }

      const name=(href)=>(map.get(href)?.name) || ipListNameCache.get(href) || href;
      const oldList=(p.oldHrefs||[]).map(h=>`<code style="font-size:11px">${escapeHtml(name(h))}</code>`).join(', ')||'<em>none</em>';
      const newList=`<code style="font-size:11px">${escapeHtml(name(p.newHref))}</code>`;

      return `<div style="margin:6px 0; padding:6px; border:1px solid #e5e7eb; border-radius:6px; background:#fff;">
        <div style="font-weight:600; margin-bottom:2px;">${title}</div>
        <div style="font-size:12px; color:#374151;">
          <div><span style="color:#6b7280;">Old:</span> ${oldList}</div>
          <div><span style="color:#6b7280;">New:</span> ${newList}</div>
          <div style="color:#6b7280;">(minimax across ${p.family?.toUpperCase?.()||'IPv4'} flows)</div>
        </div></div>`;
    };
    const parts=[];
    if (proposals.consumers) parts.push(section('CONSUMERS',proposals.consumers));
    if (proposals.providers) parts.push(section('PROVIDERS',proposals.providers));
    return parts.join('');
  }

  function replaceIpListActorWithSingle(actors,newHref){
    const out=[]; let added=false;
    for (const a of (actors||[])) { if (a?.ip_list?.href) continue; out.push(a); }
    if (newHref){ out.push({ip_list:{href:newHref}}); added=true; }
    return {actors:out, changed:added};
  }

  async function updateRuleReplaceSideWithSingle(ruleHref,newConsumerHref,newProviderHref){
    const current=await apiGet(ruleHref);
    let consumers=current.consumers??[], providers=current.providers??[], changed=false;
    if (newConsumerHref){ const res=replaceIpListActorWithSingle(consumers,newConsumerHref); consumers=res.actors; changed = changed || res.changed; }
    if (newProviderHref){ const res=replaceIpListActorWithSingle(providers,newProviderHref); providers=res.actors; changed = changed || res.changed; }
    if (!changed) return { changed:false };
    const payload={
      providers, consumers, enabled:current.enabled??true,
      ingress_services:current.ingress_services??[], egress_services:current.egress_services??[],
      network_type:current.network_type??'', description:current.description??'',
      consuming_security_principals:current.consuming_security_principals??[],
      sec_connect:current.sec_connect??false, machine_auth:current.machine_auth??false,
      stateless:current.stateless??false, unscoped_consumers:current.unscoped_consumers??false,
      use_workload_subnets:current.use_workload_subnets??[], resolve_labels_as:current.resolve_labels_as??{}
    };
    await apiPut(ruleHref,payload);
    return { changed:true };
  }

  /***************************************************************************
   * SINGLE-RULE REVIEW (row button) — Ruleset page
   ***************************************************************************/
  async function onReviewClick(btn){
    ABORT_REQUESTED=false; // reset per-run
    try {
      const rsId=getRulesetIdFromUrl(); ORG_ID=detectOrgIdFromUrlOrState()||ORG_ID;
      if (!isRulesetDetailPage()||!rsId) return toast('Not on a ruleset detail page.','warn');
      if (!ORG_ID) return toast('orgId not detected yet. Try refreshing.','warn');

      createHUD('—'); setHUDStep('start'); setHUDStatus('fetchRuleset');
      const ruleset=await fetchRuleset(rsId); setHUDStep('fetchRuleset');
      if (!ruleset){ setHUDStep('error'); toast('Failed to fetch ruleset.','err'); return; }

      setHUDStep('matchRule');
      const row=btn.closest('[data-tid="comp-grid-row"]'); if(!row){ setHUDStep('error'); toast('Could not locate rule row.','err'); return; }
      const result = await processRowCore(row, ruleset, { mode:'single' });
      if (result?.reloadRequested) setTimeout(()=>location.reload(), 900);
    } catch (e) { error(e); setHUDStep('error'); toast(`Error: ${e?.message||e}`,'err',4200); }
  }

  /**
   * Core processing for a RULE DOM row (ruleset page)
   */
  async function processRowCore(row, ruleset, { mode }){
    const useBatch = (mode === 'batch');

    // 1) Match rule for this row
    const domSig=domSignature(row);
    const rule=(ruleset.rules||[]).find(r=>matchDomToRule(domSig,r));
    if (!rule){
      if (useBatch) updateBatchHUDStatus('No matching rule for a row — skipped.');
      else { setHUDStep('error'); setHUDStatus('No matching rule found for this row.'); }
      return { skipped:true };
    }

    // SINGLE RUN: make sure HUD shows the matched rule ID immediately
    if (!useBatch) {
      const rid = rule.href?.split('/').pop() || '—';
      setHUDRuleId(rid);
    }

    return await processRuleObjectCore(rule, ruleset, { mode });
  }

  /**
   * Core processing for a RULE OBJECT (Policies & batch via API) + single-run HUD updates
   */
  async function processRuleObjectCore(rule, ruleset, { mode }) {
    const useBatch = (mode === 'batch' || mode === 'policies');
    const ruleId=rule.href?.split('/').pop()||'—';
    if (useBatch && BHUD_CURR) updateBatchHUDCurrent(`Rule ${ruleId}`);

    // 2) Rule-level pending → skip
    if (rule.update_type !== null){
      if (useBatch) { if (PHUD_STATUS) updatePoliciesHUDStatus('Rule pending update — skipped.'); }
      toast(`Rule ${ruleId}: pending → skipped`, 'warn', 2200);
      return { skipped:true };
    }

    // 3) Effective signature
    const scopeLabels=getSingleScopeLabels(ruleset);
    const raw=rawApiSignature(rule);
    const eff=effectiveApiSignature(raw,rule,scopeLabels);

    // 4) Query
    if (!useBatch) setHUDStep('createQuery');
    const result=await runTrafficQuery(eff,ruleId,rule);
    if (result.status==='aborted'){ return { cancelled:true }; }
    if (result.status!=='completed'){
      if (useBatch) { if (PHUD_STATUS) updatePoliciesHUDStatus(`Query error: ${result.status}`); }
      else { setHUDStep('error'); setHUDStatus(`Query error: ${result.status}`); toast(`Query error: ${result.status}`,'err',3600); }
      return { failed:true };
    }

    const flows=result.flowsCount??0;
    if (!useBatch) {
      setHUDStep('complete');
      setHUDFlows(flows);
      setHUDStatus(`Flows = ${flows}`);
    }

    // 5) Zero-flow: single-run = manual Disable; batch = auto-disable if setting ON
    if (flows===0){
      if (useBatch && SETTINGS.autoDisableOnZeroFlows){
        try {
          await disableRule(rule.href);
          toast(`Rule ${ruleId}: disabled (0 flows).`, 'ok', 2000);
          return { disabled:true };
        } catch (e){
          error(e);
          if (!useBatch) setHUDStep('error');
          toast(`Disable failed: ${e?.message||e}`,'err',3600);
          return { failed:true };
        }
      } else {
        // SINGLE RUN (or batch with setting OFF): present a manual Disable button
        if (!useBatch) {
          showDecisionRow(true);
          setDisableButtonEnabled(true);
          updateDisableButtonText();
          if (HUD_BTN_DISABLE){
            HUD_BTN_DISABLE.onclick = async () => {
              if (HUD_BTN_DISABLE.disabled) return;
              try {
                setDisableButtonEnabled(false);
                setHUDStatus('Disabling…');
                await disableRule(rule.href);
                setHUDStatus(SETTINGS.reloadAfterDisable ? 'Disabled. Reloading…' : 'Disabled.');
                toast(`Rule ${ruleId}: disabled.`, 'ok', 2200);
                if (SETTINGS.reloadAfterDisable) setTimeout(()=>location.reload(),900);
              } catch(e){
                setDisableButtonEnabled(true);
                setHUDStatus('Error disabling');
                error(e);
                toast(`Disable failed: ${e?.message||e}`,'err',3600);
              }
            };
          }
        }
        return { zeroFlows:true };
      }
    }

    // 6) Tightening: single-run shows proposal UI; batch/policies auto-apply
    if (!SETTINGS.tightenIpListsEnabled){
      return { flows };
    }

    try {
      const flowsJson=await downloadAsyncQueryResultsJson(result.queryHref,result.downloadPath);
      if (!Array.isArray(flowsJson) || !flowsJson.length){
        return { flows };
      }

      const { v4Flows, v6Flows } = partitionFlowsByFamily(flowsJson);
      const {consumerIpLists, providerIpLists} = extractIpListHrefsFromRule(rule);

      // Consumer family
      const v4SrcSet = new Set(); for (const f of v4Flows) (f?.src?.ip_lists||[]).forEach(l=>{ if (l?.href) v4SrcSet.add(l.href); });
      const v6SrcSet = new Set(); for (const f of v6Flows) (f?.src?.ip_lists||[]).forEach(l=>{ if (l?.href) v6SrcSet.add(l.href); });
      let consumerFamily = 'v4';
      if (consumerIpLists.some(h=>v4SrcSet.has(h))) consumerFamily = 'v4';
      else if (consumerIpLists.some(h=>v6SrcSet.has(h))) consumerFamily = 'v6';
      else if (!v4Flows.length && v6Flows.length) consumerFamily = 'v6';

      // Provider family detection (fix kept)
      let providerFamily = 'v4';
      const v4DstSet = new Set(); for (const f of v4Flows) (f?.dst?.ip_lists || []).forEach(l => { if (l?.href) v4DstSet.add(l.href); });
      const v6DstSet = new Set(); for (const f of v6Flows) (f?.dst?.ip_lists || []).forEach(l => { if (l?.href) v6DstSet.add(l.href); });
      if (providerIpLists.some(h => v4DstSet.has(h))) providerFamily = 'v4';
      else if (providerIpLists.some(h => v6DstSet.has(h))) providerFamily = 'v6';
      else if (!v4Flows.length && v6Flows.length) providerFamily = 'v6';

      let chosenConsumer = null, consumerNameMap = null;
      if (consumerIpLists.length) {
        const famFlows = consumerFamily === 'v6' ? v6Flows : v4Flows;
        const idxStats = buildIndexStats(famFlows, 'src');
        const common   = computeCommonHrefSet(famFlows, 'src');
        consumerNameMap = idxStats;
        chosenConsumer  = common.common.size ? chooseMinimaxFromSet(common.common, idxStats) : null;
      }

      let chosenProvider = null, providerNameMap = null;
      if (providerIpLists.length) {
        const famFlows = providerFamily === 'v6' ? v6Flows : v4Flows;
        const idxStats = buildIndexStats(famFlows, 'dst');
        const common   = computeCommonHrefSet(famFlows, 'dst');
        providerNameMap = idxStats;
        chosenProvider  = common.common.size ? chooseMinimaxFromSet(common.common, idxStats) : null;
      }

      const proposals={};
      if (consumerIpLists.length && chosenConsumer) {
        const same=(consumerIpLists.length===1 && consumerIpLists[0]===chosenConsumer);
        if (!same) proposals.consumers={oldHrefs:consumerIpLists,newHref:chosenConsumer,nameMap:consumerNameMap,family:consumerFamily};
      }
      if (providerIpLists.length && chosenProvider) {
        const same=(providerIpLists.length===1 && providerIpLists[0]===chosenProvider);
        if (!same) proposals.providers={oldHrefs:providerIpLists,newHref:chosenProvider,nameMap:providerNameMap,family:providerFamily};
      }

      const hasProposal = !!(proposals.consumers || proposals.providers);

      if (useBatch) {
        // AUTO-APPLY in batch/policies if proposals exist
        if (hasProposal){
          try {
            const consumerHref=proposals.consumers?.newHref??null, providerHref=proposals.providers?.newHref??null;
            const upd = await updateRuleReplaceSideWithSingle(rule.href,consumerHref,providerHref);
            return {
              flows,
              tightened: !!upd.changed,
              tightenedDetail: {
                ruleId,
                consumerOld: proposals.consumers?.oldHrefs||[],
                consumerNew: proposals.consumers?.newHref||null,
                providerOld: proposals.providers?.oldHrefs||[],
                providerNew: proposals.providers?.newHref||null
              }
            };
          } catch (e) {
            error(e);
            return { flows };
          }
        }
        return { flows };
      }

      // SINGLE-RUN: render proposal details with Apply / Skip
      const hrefs = new Set();
      if (proposals.consumers){ proposals.consumers.oldHrefs?.forEach(h=>hrefs.add(h)); if (proposals.consumers.newHref) hrefs.add(proposals.consumers.newHref); }
      if (proposals.providers){ proposals.providers.oldHrefs?.forEach(h=>hrefs.add(h)); if (proposals.providers.newHref) hrefs.add(proposals.providers.newHref); }
      await ensureIpListNames([...hrefs]);

      if (hasProposal) {
        const html = renderTightenProposalHtml(proposals);
        showTightenPanel(html, true);

        if (HUD_TIGHTEN_APPLY){
          HUD_TIGHTEN_APPLY.disabled=false;
          HUD_TIGHTEN_APPLY.onclick=async()=>{ try{
            HUD_TIGHTEN_APPLY.disabled=true; if (HUD_TIGHTEN_SKIP) HUD_TIGHTEN_SKIP.disabled=true;
            setHUDStatus('Applying tightening…');
            const consumerHref=proposals.consumers?.newHref??null, providerHref=proposals.providers?.newHref??null;
            await updateRuleReplaceSideWithSingle(rule.href,consumerHref,providerHref);
            setHUDStatus('Applied. Reloading…'); toast(`Rule ${ruleId}: tightening applied.`, 'ok', 2500);
            setTimeout(()=>location.reload(),900);
          } catch(e){ HUD_TIGHTEN_APPLY.disabled=false; if (HUD_TIGHTEN_SKIP) HUD_TIGHTEN_SKIP.disabled=false; setHUDStatus('Error applying tightening (see console).'); error(e); toast(`Tighten failed: ${e?.message||e}`,'err',4200); } };
        }
        if (HUD_TIGHTEN_SKIP){
          HUD_TIGHTEN_SKIP.disabled=false;
          HUD_TIGHTEN_SKIP.onclick=()=>{ showTightenPanel('',false); setHUDStatus(`Flows = ${flows}. Proposal skipped.`); };
        }
        return { flows, proposal:true };
      } else {
        // No tightening opportunity (flows present, but no change recommended)
        showTightenPanel('', false);
        if (!useBatch) {
          const msg = 'Flows detected, but no tightening opportunity found.';
          setHUDStatus(msg);
          toast(msg, 'ok', 2200);
        }
        return { flows, proposal:false };
      }

    } catch (e) {
      error(e);
      return { flows };
    }
  }

  /***************************************************************************
   * Row button injection (clone Edit → Review) — Ruleset page
   ***************************************************************************/
  const GRID_SELECTOR='div[data-tid~="comp-grid"][data-tid~="comp-grid-allow"]';
  const ROW_SELECTOR='div[data-tid="comp-grid-row"]';
  const BUTTONS_CELL_SELECTOR='div[data-tid="comp-grid-column-buttons"]';
  const ROW_BUTTONS_WRAP_SELECTOR='.Qu';
  const EDIT_BUTTON_SELECTOR='button[data-tid~="comp-button"][data-tid~="comp-button-edit"], button[aria-label="Edit"]';
  const REVIEW_LABEL='Review';
  const REVIEW_BTN_DATA_TID='comp-button comp-button-review';
  const REVIEW_ICON_DATA_TID='comp-icon comp-icon-search';
  const ICON_IDS=['#search','#magnify','#magnifier','#zoom'];

  function hasReviewButton(container){ return !!container.querySelector('button[data-userscript-review="true"]'); }
  function pickIconId(){ for (const id of ICON_IDS){ const clean=id.slice(1); if (document.querySelector(`symbol#${clean}`)) return id; if (document.querySelector(`use[href="${id}"], use[xlink\\:href="${id}"]`)) return id; } return ICON_IDS[0]; }
  function buildReviewButtonFrom(editBtn){
    const btn=editBtn.cloneNode(true);
    btn.setAttribute('data-userscript-review','true'); btn.setAttribute('aria-label',REVIEW_LABEL); btn.setAttribute('title',REVIEW_LABEL);
    const dtid=btn.getAttribute('data-tid')||'comp-button';
    const tokens=dtid.split(/\s+/).filter(Boolean).map(t=>t.replace(/comp-button-edit/gi,'comp-button-review'));
    if (!tokens.includes('comp-button-review')) tokens.push('comp-button-review'); btn.setAttribute('data-tid',tokens.join(' '));
    const outer=btn.querySelector(':scope > div')||btn;

    const iconSpan=outer.querySelector('span[data-tid*="comp-icon"]')||outer.querySelector('span');
    if (iconSpan){
      iconSpan.setAttribute('aria-label','Search'); iconSpan.setAttribute('data-tid',REVIEW_ICON_DATA_TID);
      const svg=iconSpan.querySelector('svg')||(()=>{ const s=document.createElementNS('http://www.w3.org/2000/svg','svg'); s.setAttribute('class','Zs'); iconSpan.appendChild(s); return s; })();
      let use=svg.querySelector('use'); if(!use){ use=document.createElementNS('http://www.w3.org/2000/svg','use'); svg.appendChild(use); }
      const iconId=pickIconId();
      use.setAttributeNS('http://www.w3.org/1999/xlink','xlink:href',iconId); use.setAttribute('href',iconId);
    }
    const label=outer.querySelector('[data-tid="button-text"]'); if (label) label.textContent='';

    btn.addEventListener('click',(e)=>{ e.stopPropagation(); onReviewClick(btn); });
    return btn;
  }
  function processRow(row){
    const buttonsCell=row.querySelector(BUTTONS_CELL_SELECTOR); if(!buttonsCell) return;
    const wrap=buttonsCell.querySelector(ROW_BUTTONS_WRAP_SELECTOR)||buttonsCell; if (hasReviewButton(wrap)) return;
    const editBtn=wrap.querySelector(EDIT_BUTTON_SELECTOR); let reviewBtn;
    if (editBtn){ reviewBtn=buildReviewButtonFrom(editBtn); editBtn.insertAdjacentElement('afterend',reviewBtn); }
    else {
      const anyBtn=wrap.querySelector('button');
      if (anyBtn){ reviewBtn=buildReviewButtonFrom(anyBtn); wrap.insertBefore(reviewBtn,wrap.firstChild); }
      else {
        const btn=document.createElement('button'); btn.type='button'; btn.setAttribute('data-userscript-review','true'); btn.setAttribute('data-tid',REVIEW_BTN_DATA_TID);
        btn.setAttribute('aria-label',REVIEW_LABEL); btn.setAttribute('title',REVIEW_LABEL);
        const outer=document.createElement('div'); outer.className='hy Bz Bw hw hx';
        const inner=document.createElement('div'); inner.className='iD';
        const iconSpan=document.createElement('span'); iconSpan.className='Zn'; iconSpan.setAttribute('data-tid',REVIEW_ICON_DATA_TID);
        const svg=document.createElementNS('http://www.w3.org/2000/svg','svg'); svg.setAttribute('class','Zs');
        const use=document.createElementNS('http://www.w3.org/2000/svg','use'); const iconId=pickIconId();
        use.setAttributeNS('http://www.w3.org/1999/xlink','xlink:href',iconId); use.setAttribute('href',iconId);
        svg.appendChild(use); iconSpan.appendChild(svg); inner.appendChild(iconSpan); outer.appendChild(inner); btn.appendChild(outer);
        btn.addEventListener('click',(e)=>{ e.stopPropagation(); onReviewClick(btn); }); wrap.appendChild(btn);
      }
    }
  }
  function processGrid(grid){ grid.querySelectorAll(ROW_SELECTOR).forEach(processRow); }
  function observeGrid(grid){
    if (grid.__userscriptReviewObserved) return; grid.__userscriptReviewObserved=true;
    const obs=new MutationObserver((muts)=>{ for (const m of muts){ for (const node of m.addedNodes){ if(!(node instanceof Element)) continue; if (node.matches?.(ROW_SELECTOR)) processRow(node); else node.querySelectorAll?.(ROW_SELECTOR)?.forEach(processRow); } } });
    obs.observe(grid,{childList:true,subtree:true});
  }
  function findTargetGrids(){ return Array.from(document.querySelectorAll(GRID_SELECTOR)); }
  function ensureRowButtons(){ if (!isRulesetDetailPage()) return false; const grids=findTargetGrids(); if (!grids.length) return false; grids.forEach((g)=>{ processGrid(g); observeGrid(g); }); return true; }

  /***************************************************************************
   * TOP TOOLBAR "REVIEW" BUTTON (Ruleset + Policies)
   ***************************************************************************/
  const TOOLBAR_SELECTOR       = 'div[data-tid="comp-toolbar"]';
  const TOOLGROUP_SELECTOR     = 'div[data-tid="comp-toolgroup"]';
  const REFRESH_BTN_SELECTOR   = 'button[data-tid~="comp-button"][data-tid~="comp-button-refresh"], button[aria-label="Refresh"]';

  const RULESET_MARKERS = [
    'nav[data-tid~="comp-menu-ruleset-actions"]',
    'span[data-tid="scope-title"]'
  ];
  const RULESET_POLICY_ACTIONS_SELECTOR = 'nav[data-tid~="comp-menu-ruleset-actions"]';

  const POLICIES_MARKERS = [
    'nav[data-tid~="comp-menu-add"]',
    'a[data-tid~="comp-button-policygenerator"]',
    'button[data-tid~="comp-button-provision"]'
  ];
  // Policies grid selectors
  const POLICIES_GRID_SELECTOR = 'div[data-tid~="comp-grid"]';
  const POLICIES_ROW_SELECTOR  = 'div[data-tid="comp-grid-row"]';
  const ROW_CHECKBOX_SELECTOR  = 'input[type="checkbox"][data-tid="elem-input"]';

  const REVIEW_ITEM_ID          = 'userscript-review-toolitem';
  const REVIEW_BTN_ID           = 'userscript-review-button';
  const REVIEW_TOOLBAR_LABEL    = 'Review';

  function inShadowRoot(node) {
    try { const root = node && node.getRootNode && node.getRootNode(); return root && root.host && root instanceof ShadowRoot; } catch { return false; }
  }
  function getToolbar(doc = document) {
    const toolbar = doc.querySelector(TOOLBAR_SELECTOR);
    if (!toolbar) return null;
    if (inShadowRoot(toolbar)) return null;
    return toolbar;
  }
  function hasAny(root, selectors) { return selectors.some(sel => root.querySelector(sel)); }

  function getPageContext() {
    const toolbar = getToolbar();
    if (!toolbar) return 'unknown';
    if (hasAny(toolbar, RULESET_MARKERS)) return 'ruleset';
    if (hasAny(toolbar, POLICIES_MARKERS)) return 'policies';
    return 'unknown';
  }
  function findTargetToolgroup(context) {
    const toolbar = getToolbar();
    if (!toolbar) return null;
    const groups = Array.from(toolbar.querySelectorAll(TOOLGROUP_SELECTOR));
    if (!groups.length) return null;

    if (context === 'ruleset') {
      const matches = groups.filter(g => g.querySelector(RULESET_POLICY_ACTIONS_SELECTOR) && g.querySelector(REFRESH_BTN_SELECTOR));
      if (matches.length) return matches[matches.length - 1];
      if (groups.length >= 2) return groups[1];
      return groups[groups.length - 1];
    }
    if (context === 'policies') {
      const withRefresh = groups.filter(g => g.querySelector(REFRESH_BTN_SELECTOR));
      if (withRefresh.length) return withRefresh[withRefresh.length - 1];
      if (groups.length >= 2) return groups[1];
      return groups[groups.length - 1];
    }
    return null;
  }
  function buildReviewToolbarItemFrom(refreshBtn, context) {
    const item = document.createElement('div');
    item.className = 'lt';
    item.setAttribute('data-tid', 'elem-toolgroup-item');
    item.id = REVIEW_ITEM_ID;

    const btn = refreshBtn.cloneNode(true);
    btn.id = REVIEW_BTN_ID;

    btn.setAttribute('aria-label', REVIEW_TOOLBAR_LABEL);
    btn.setAttribute('title', REVIEW_TOOLBAR_LABEL);
    btn.dataset.pageContext = context;

    const dtid = btn.getAttribute('data-tid') || 'comp-button';
    const tokens = dtid.split(/\s+/).filter(Boolean).map(t => t.replace(/comp-button-refresh/gi, 'comp-button-review'));
    if (!tokens.includes('comp-button-review')) tokens.push('comp-button-review');
    btn.setAttribute('data-tid', tokens.join(' '));

    const outerWrap = btn.querySelector(':scope > div') || btn;

    const labelSpan = outerWrap.querySelector(':scope [data-tid="button-text"]') || outerWrap.querySelector(':scope span');
    if (labelSpan) labelSpan.textContent = REVIEW_TOOLBAR_LABEL;

    const iconSpan = outerWrap.querySelector(':scope span[data-tid*="comp-icon"]') || outerWrap.querySelector(':scope span');
    if (iconSpan) {
      iconSpan.setAttribute('aria-label', 'Search');
      iconSpan.setAttribute('data-tid', 'comp-icon comp-icon-search');
      let svg = iconSpan.querySelector('svg');
      if (!svg) { svg = document.createElementNS('http://www.w3.org/2000/svg','svg'); svg.setAttribute('class','Zs'); iconSpan.appendChild(svg); }
      let use = svg.querySelector('use'); if (!use){ use=document.createElementNS('http://www.w3.org/2000/svg','use'); svg.appendChild(use); }
      const SEARCH_SYMBOL_ID = '#search';
      use.setAttributeNS('http://www.w3.org/1999/xlink','xlink:href', SEARCH_SYMBOL_ID);
      use.setAttribute('href', SEARCH_SYMBOL_ID);
    }

    btn.onclick = async (e) => {
      e.stopPropagation();
      const ctx = btn.dataset.pageContext || getPageContext();
      if (ctx === 'policies') onToolbarPoliciesReviewClick();
      else await onToolbarRulesetReviewAllClick();
    };

    item.appendChild(btn);
    return item;
  }
  function ensureToolbarReviewItem() {
    const context = getPageContext();
    if (context === 'unknown') return null;

    const toolgroup = findTargetToolgroup(context);
    if (!toolgroup) return null;

    let existing = document.getElementById(REVIEW_ITEM_ID);
    if (existing) {
      if (existing.parentElement !== toolgroup) toolgroup.appendChild(existing);
      if (toolgroup.lastElementChild !== existing) toolgroup.appendChild(existing);
      const btn = existing.querySelector(`#${REVIEW_BTN_ID}`) || existing.querySelector('button');
      if (btn) {
        btn.dataset.pageContext = context;
        btn.onclick = async (e) => {
          e.stopPropagation();
          const ctx = btn.dataset.pageContext || getPageContext();
          if (ctx === 'policies') onToolbarPoliciesReviewClick();
          else await onToolbarRulesetReviewAllClick();
        };
      }
      return existing;
    }

    const refreshBtn = toolgroup.querySelector(REFRESH_BTN_SELECTOR);
    if (!refreshBtn) return null;

    const reviewItem = buildReviewToolbarItemFrom(refreshBtn, context);
    toolgroup.appendChild(reviewItem);
    return reviewItem;
  }

  /***************************************************************************
   * POLICIES PAGE — ON-PAGE BATCH (no navigation)
   ***************************************************************************/
  function getPoliciesGridRoot() {
    const roots = Array.from(document.querySelectorAll(POLICIES_GRID_SELECTOR));
    return roots.find(r => r.querySelector(POLICIES_ROW_SELECTOR)) || null;
  }
  function getPoliciesRows() {
    const root = getPoliciesGridRoot();
    if (!root) return [];
    return Array.from(root.querySelectorAll(POLICIES_ROW_SELECTOR));
  }
  function getRowText(row, tid) {
    return (row.querySelector(`[data-tid="${tid}"]`)?.textContent || '').trim();
  }
  function extractRulesetRowKey(row) {
    const a = row.querySelector('a[href*="#/rulesets/"]');
    const href = a ? (a.getAttribute('href') || a.href || '') : '';
    const idMatch = href.match(/#\/rulesets\/(\d+)/);
    const id = idMatch ? idMatch[1] : null;

    const name = getRowText(row, 'comp-grid-column-name') || '';
    const updatedAt = getRowText(row, 'comp-grid-column-updatedat') || '';

    return { id, name, updatedAt };
  }
  function rowsChecked(rows) {
    return rows.filter(r => !!r.querySelector(`${ROW_CHECKBOX_SELECTOR}:checked`));
  }

  async function resolveRulesetIdsFromPoliciesSelection() {
    const rows = rowsChecked(getPoliciesRows());
    const selected = rows.map(extractRulesetRowKey);
    const needLookup = selected.filter(s => !s.id);

    if (!needLookup.length) return selected.map(s => ({ id: s.id, name: s.name, updatedAt: s.updatedAt }));

    // Fallback: list rulesets and map by name + updatedAt (best effort)
    try {
      const list = await listRulesets(5000);
      const items = Array.isArray(list?.items) ? list.items : Array.isArray(list) ? list : [];
      const mapped = selected.map(s => {
        if (s.id) return { id: s.id, name: s.name, updatedAt: s.updatedAt };
        const nameMatches = items.filter(it => (it?.name || '').trim() === s.name.trim());
        if (!nameMatches.length) return { id: null, name: s.name, updatedAt: s.updatedAt };
        if (nameMatches.length === 1) {
          const href = nameMatches[0]?.href || '';
          const idm = href.match(/\/rule_sets\/(\d+)/);
          return { id: idm ? idm[1] : null, name: s.name, updatedAt: s.updatedAt };
        }
        let displayTime = Date.parse(s.updatedAt);
        if (!Number.isFinite(displayTime)) displayTime = 0;
        let best = null, bestDiff = Number.POSITIVE_INFINITY;
        for (const it of nameMatches) {
          const href = it?.href || '';
          const idm = href.match(/\/rule_sets\/(\d+)/);
          const id = idm ? idm[1] : null;
          const apiTime = Date.parse(it?.updated_at || it?.update_time || '');
          const diff = Number.isFinite(apiTime) && Number.isFinite(displayTime) ? Math.abs(apiTime - displayTime) : 0;
          if (diff < bestDiff) { bestDiff = diff; best = { id, name: s.name, updatedAt: s.updatedAt }; }
        }
        return best || { id: null, name: s.name, updatedAt: s.updatedAt };
      });
      return mapped;
    } catch (e) {
      warn('Ruleset list lookup failed; unresolved IDs will be skipped.', e);
      return selected.map(s => ({ id: s.id, name: s.name, updatedAt: s.updatedAt }));
    }
  }

  async function onToolbarPoliciesReviewClick() {
    try {
      ORG_ID=detectOrgIdFromUrlOrState()||ORG_ID;
      if (!ORG_ID) return toast('orgId not detected yet. Try refreshing.','warn');

      const resolved = await resolveRulesetIdsFromPoliciesSelection();
      const targets = resolved.filter(x => !!x.id);
      const skippedNoId = resolved.filter(x => !x.id);

      if (skippedNoId.length) toast(`${skippedNoId.length} selected ruleset(s) could not resolve ID and will be skipped.`, 'warn', 3600);
      if (!targets.length) {
        toast('No resolvable selected rulesets on Policies page.', 'warn', 2400);
        return;
      }

      // START on-page policies HUD run
      POLICIES_ABORT_REQUESTED=false;
      createPoliciesHUD(targets.length);
      updatePoliciesHUDStatus('Starting…');

      let rsProcessed=0, rsSkipped=0, rsFailed=0;
      let rulesProcessed=0, rulesSkipped=0, rulesFailed=0, rulesDisabled=0, rulesTightened=0;

      for (const t of targets) {
        if (POLICIES_ABORT_REQUESTED) { updatePoliciesHUDStatus('Cancelled.'); break; }

        const rsLabel = `${t.name || 'Ruleset'} (#${t.id})`;
        updatePoliciesHUDCurrentRS(rsLabel);
        updatePoliciesHUDCurrentRule('—');

        // Fetch ruleset
        let ruleset=null;
        try {
          ruleset = await fetchRuleset(t.id);
        } catch (e) {
          rsFailed++;
          updatePoliciesHUDStatus(`Failed to fetch ${rsLabel}. Skipping…`);
          updatePoliciesHUDCounters({rsProcessed, rsTotal:targets.length, rsSkipped, rsFailed, rulesProcessed, rulesSkipped, rulesFailed, rulesDisabled, rulesTightened});
          continue;
        }

        if (!ruleset) {
          rsFailed++;
          updatePoliciesHUDStatus(`Empty response for ${rsLabel}. Skipping…`);
          updatePoliciesHUDCounters({rsProcessed, rsTotal:targets.length, rsSkipped, rsFailed, rulesProcessed, rulesSkipped, rulesFailed, rulesDisabled, rulesTightened});
          continue;
        }

        if (ruleset.update_type !== null) {
          rsSkipped++;
          appendPoliciesDetail(`<strong>Ruleset skipped (pending update):</strong> <code>${escapeHtml(rsLabel)}</code>`);
          updatePoliciesHUDCounters({rsProcessed, rsTotal:targets.length, rsSkipped, rsFailed, rulesProcessed, rulesSkipped, rulesFailed, rulesDisabled, rulesTightened});
          continue;
        }

        const rules = Array.isArray(ruleset.rules) ? ruleset.rules : [];
        if (!rules.length) {
          rsProcessed++; // processed but 0 rules
          appendPoliciesDetail(`<strong>Ruleset processed:</strong> <code>${escapeHtml(rsLabel)}</code> · 0 rules`);
          updatePoliciesHUDCounters({rsProcessed, rsTotal:targets.length, rsSkipped, rsFailed, rulesProcessed, rulesSkipped, rulesFailed, rulesDisabled, rulesTightened});
          continue;
        }

        // Process rules sequentially — auto actions enabled here
        let localProcessed=0, localSkipped=0, localFailed=0, localDisabled=0, localTightened=0;

        for (const rule of rules) {
          if (POLICIES_ABORT_REQUESTED) break;

          const rid=rule.href?.split('/').pop()||'—';
          updatePoliciesHUDCurrentRule(`Rule ${rid}`);

          try {
            const res = await processRuleObjectCore(rule, ruleset, { mode:'policies' });
            localProcessed++; rulesProcessed++;
            if (res?.skipped) { localSkipped++; rulesSkipped++; }
            if (res?.failed)  { localFailed++;  rulesFailed++; }
            if (res?.disabled){ localDisabled++; rulesDisabled++; }
            if (res?.tightened){ localTightened++; rulesTightened++; }
          } catch (e) {
            localFailed++; rulesFailed++;
            error(e);
          }

          updatePoliciesHUDCounters({rsProcessed, rsTotal:targets.length, rsSkipped, rsFailed, rulesProcessed, rulesSkipped, rulesFailed, rulesDisabled, rulesTightened});
        }

        rsProcessed++;
        appendPoliciesDetail(
          `<strong>Ruleset processed:</strong> <code>${escapeHtml(rsLabel)}</code>` +
          `<div style="margin-top:4px">Rules: ${localProcessed} processed · ${localSkipped} skipped · ${localFailed} failed · ${localDisabled} disabled · ${localTightened} tightened</div>`
        );
        updatePoliciesHUDCounters({rsProcessed, rsTotal:targets.length, rsSkipped, rsFailed, rulesProcessed, rulesSkipped, rulesFailed, rulesDisabled, rulesTightened});
      }

      updatePoliciesHUDCurrentRS('—');
      updatePoliciesHUDCurrentRule('—');
      updatePoliciesHUDStatus('Completed.');

      const summary = `Rulesets: ${rsProcessed}/${targets.length} processed, ${rsSkipped} skipped, ${rsFailed} failed. Rules: ${rulesProcessed} processed · ${rulesSkipped} skipped · ${rulesFailed} failed · ${rulesDisabled} disabled · ${rulesTightened} tightened.`;
      toast(summary, 'ok', 3000);

      // Reload Policies only if changes were applied and user didn’t cancel
      const changesApplied = (rulesDisabled + rulesTightened) > 0;
      if (!POLICIES_ABORT_REQUESTED && changesApplied) {
        toast(`${rulesDisabled + rulesTightened} change(s) applied. Reloading Policies to reflect updates…`, 'ok', 1600);
        setTimeout(() => location.reload(), 1000);
      } else if (!changesApplied) {
        toast('No rule changes detected — no reload needed.', 'ok', 1800);
      }
    } catch (e) {
      error(e);
      toast(`Policies on-page batch failed: ${e?.message || e}`, 'err', 4200);
    }
  }

  /***************************************************************************
   * RULESET TOOLBAR CLICK HANDLER — Batch
   ***************************************************************************/
  function getSelectedRuleRows() {
    const grid = document.querySelector(GRID_SELECTOR);
    if (!grid) return [];
    const all = Array.from(grid.querySelectorAll(ROW_SELECTOR));
    return all.filter(row => !!row.querySelector(`${ROW_CHECKBOX_SELECTOR}:checked`));
  }

  async function onToolbarRulesetReviewAllClick(opts = {}) {
    const suppressReload = !!opts.suppressReload;
    const onlySelected   = !!opts.onlySelected;
    try {
      const rsId=getRulesetIdFromUrl(); ORG_ID=detectOrgIdFromUrlOrState()||ORG_ID;
      if (!isRulesetDetailPage()||!rsId) return toast('Not on a ruleset detail page.','warn');
      if (!ORG_ID) return toast('orgId not detected yet. Try refreshing.','warn');

      const ruleset = await fetchRuleset(rsId);
      if (!ruleset) { toast('Failed to fetch ruleset.','err'); return; }
      if (ruleset.update_type !== null) {
        toast('Ruleset has a pending update — Review skipped.', 'warn', 4200);
        return { processed: 0, skipped: 1, failed: 0, disabled: 0, tightened: 0 };
      }

      ensureRowButtons();

      let rows = onlySelected ? getSelectedRuleRows()
                              : Array.from(document.querySelectorAll(`${GRID_SELECTOR} ${ROW_SELECTOR}`));
      if (onlySelected && rows.length === 0) {
        toast('No rules selected in this ruleset — skipping.', 'warn', 2200);
        return { processed: 0, skipped: 1, failed: 0, disabled: 0, tightened: 0 };
      }
      if (!rows.length) { toast('No rules found to review.', 'warn', 2200); return; }

      BATCH_MODE = true; BATCH_ABORT_REQUESTED=false;
      createBatchHUD(rows.length);
      updateBatchHUDStatus('Starting…');

      let processed=0, skipped=0, failed=0, disabled=0, tightened=0;
      let anyChanges=false;
      const tightenRecords=[]; // {ruleId, consumerOld[], consumerNew, providerOld[], providerNew}

      for (const row of rows) {
        if (BATCH_ABORT_REQUESTED) { updateBatchHUDStatus('Batch cancelled.'); break; }
        processRow(row); // ensure row has a Review button (DOM-only)

        const res = await processRowCore(row, ruleset, { mode:'batch' });

        processed++;
        if (res?.skipped) skipped++;
        if (res?.failed)  failed++;
        if (res?.disabled){ disabled++; anyChanges=true; appendBatchDetail(`<strong>Disabled:</strong> Rule <code>${escapeHtml(res.ruleId || '')}</code>`); }
        if (res?.tightened){
          tightened++; anyChanges=true;
          if (res.tightenedDetail) tightenRecords.push(res.tightenedDetail);
        }

        updateBatchHUDCounters({processed, total: rows.length, skipped, failed, disabled, tightened});
      }

      // Resolve names for summary once (batch-friendly)
      const summaryHrefs=new Set();
      for (const rec of tightenRecords){
        (rec.consumerOld||[]).forEach(h=>summaryHrefs.add(h));
        (rec.providerOld||[]).forEach(h=>summaryHrefs.add(h));
        if (rec.consumerNew) summaryHrefs.add(rec.consumerNew);
        if (rec.providerNew) summaryHrefs.add(rec.providerNew);
      }
      if (summaryHrefs.size) await ensureIpListNames([...summaryHrefs]);

      // Render tighten summary
      for (const rec of tightenRecords){
        const name=(h)=> ipListNameCache.get(h)||h;
        const parts=[];
        if ((rec.consumerOld||[]).length || rec.consumerNew){
          const old= (rec.consumerOld||[]).map(h=>`<code>${escapeHtml(name(h))}</code>`).join(', ')||'<em>none</em>';
          const neu= rec.consumerNew ? `<code>${escapeHtml(name(rec.consumerNew))}</code>` : '<em>none</em>';
          parts.push(`<div>Consumers: ${old} → ${neu}</div>`);
        }
        if ((rec.providerOld||[]).length || rec.providerNew){
          const old= (rec.providerOld||[]).map(h=>`<code>${escapeHtml(name(h))}</code>`).join(', ')||'<em>none</em>';
          const neu= rec.providerNew ? `<code>${escapeHtml(name(rec.providerNew))}</code>` : '<em>none</em>';
          parts.push(`<div>Providers: ${old} → ${neu}</div>`);
        }
        appendBatchDetail(`<strong>Tightened:</strong> Rule <code>${escapeHtml(rec.ruleId)}</code><div style="margin-top:4px">${parts.join('')}</div>`);
      }

      updateBatchHUDStatus('Completed.');
      updateBatchHUDCurrent('—');

      if (anyChanges && !suppressReload){
        toast('Reloading to reflect changes…', 'ok', 1600);
        setTimeout(()=>location.reload(), 1000);
      }

      return { processed, skipped, failed, disabled, tightened };

    } catch (e) {
      error(e);
      toast(`Batch failed: ${e?.message||e}`, 'err', 4200);
    } finally {
      BATCH_MODE = false;
    }
  }

  /***************************************************************************
   * INIT
   ***************************************************************************/
  function init(){
    ensureRowButtons();
    ensureToolbarReviewItem();

    const globalObs=new MutationObserver(()=>{
      ensureRowButtons();
      ensureToolbarReviewItem();
    });
    globalObs.observe(document.documentElement,{childList:true,subtree:true});

    let attempts=0;
    const timer=setInterval(()=>{
      attempts++;
      const done = ensureRowButtons();
      ensureToolbarReviewItem();
      if (done || attempts>=30) clearInterval(timer);
    },300);

    window.addEventListener('hashchange',()=>setTimeout(()=>{
      ensureRowButtons();
      ensureToolbarReviewItem();
    },50));
  }
  if (document.readyState==='complete' || document.readyState==='interactive') init();
  else window.addEventListener('DOMContentLoaded', init);

})();