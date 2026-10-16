/* OKV SACCO — Subscribe (Online) — shared configuration.
   Paste your deployed Apps Script Web App URL below (ends in /exec). */
const API_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";

/* Calls the backend. Uses no explicit Content-Type header so the browser
   sends it as text/plain — this avoids a CORS preflight request, which
   Apps Script web apps do not handle. Code.gs still parses it as JSON. */
async function apiCall(action, payload){
  const res = await fetch(API_URL, {
    method: "POST",
    body: JSON.stringify(Object.assign({ action }, payload || {}))
  });
  const data = await res.json();
  if(!data.ok) throw new Error(data.error || "Request failed.");
  return data;
}
async function sha256(text){
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(buf)).map(b=> b.toString(16).padStart(2,"0")).join("");
}

/* Renewal contact details shown on the Admin's Subscription tab and the
   marketing pages. These are FALLBACK values only, shown instantly so pages
   don't flash empty — loadPlatformSettings() below overwrites them with the
   live values from the Settings sheet, which the Super Admin can edit from
   the Settings tab on their dashboard. */
const CONTACT_INFO = {
  whatsapp: "+2348104141138",
  phone: "+2348104141138",
  email: "technologyokv@gmail.com",
  website: "www.okvtechnology.com"
};

/* Bank transfer details shown on the Upgrade/Subscription form.
   Fallback only — see the note on CONTACT_INFO above. */
const BANK_DETAILS = { accountNumber:"8104141138", bankName:"OPay MFB", accountName:"Olasile Kehinde Victor" };

/* Fetches the live contact & bank details from the backend and mutates
   CONTACT_INFO/BANK_DETAILS in place (their bindings stay const, only their
   contents change), so every page that reads CONTACT_INFO.x / BANK_DETAILS.x
   picks up the update automatically as long as this runs before that read.
   Safe to call with no session — falls back silently (keeping the defaults
   above) if offline or the request fails. */
let PLATFORM_SETTINGS = null; // last full getSettings() response, for pages that need more than contact/bank
async function loadPlatformSettings(){
  try{
    const result = await apiCall("getSettings", {});
    PLATFORM_SETTINGS = result;
    if(result.contact) Object.assign(CONTACT_INFO, result.contact);
    if(result.bank) Object.assign(BANK_DETAILS, result.bank);
  }catch(e){ /* keep fallback values */ }
}

/* Applies the Super Admin's branding (colors, logo, system name/tagline) to
   the current page. Call this after loadPlatformSettings(). Expects, where
   present: #brandMark (the default "S" logo block), #brandLogoImg (an <img>
   hidden by default), and any number of elements with class "brand-name"
   or "brand-tagline" whose text should reflect the configured name/tagline. */
function applyBranding(){
  const b = PLATFORM_SETTINGS && PLATFORM_SETTINGS.branding;
  if(!b) return;
  if(b.primaryColor) document.documentElement.style.setProperty("--teal-2", b.primaryColor);
  if(b.secondaryColor){
    document.documentElement.style.setProperty("--navy", b.secondaryColor);
  }
  if(b.logoUrl){
    const mark = document.getElementById("brandMark");
    const img = document.getElementById("brandLogoImg");
    if(img){ img.src = b.logoUrl; img.style.display = ""; }
    if(mark) mark.style.display = "none";
  }
  if(b.systemName){
    document.querySelectorAll(".brand-name").forEach(el=> el.textContent = b.systemName);
    document.title = document.title.replace(/OKV SACCO/g, b.systemName);
  }
  if(b.tagline){
    document.querySelectorAll(".brand-tagline").forEach(el=> el.textContent = b.tagline);
  }
}


/* Plan catalogue fallback — used only until loadPlans() succeeds, so pages
   never render empty while the live catalogue (Super Admin > Plans &
   Pricing) loads. Every price is in NGN. "id" must stay unique and stable
   since it's what gets stored on the Users sheet as each org's chosen plan. */
const PLANS = {
  monthly: [
    { id:"basic-monthly", name:"Basic", cycle:"Monthly", price:15000,
      audience:"Small cooperatives and single-branch SACCOs just getting started",
      features:["Up to 200 members","Core registers: Members, Contributions, Loans, Repayments, Dividends","2 staff logins","Member communication (Email/SMS/WhatsApp)","Downloadable role-based user manual","Email support"] },
    { id:"professional-monthly", name:"Professional", cycle:"Monthly", price:35000,
      audience:"Growing SACCOs and multi-officer teams that need full team access",
      features:["Unlimited members","Everything in Basic","Unlimited staff logins & role assignment","Priority support","Data export"] }
  ],
  biannual: [
    { id:"basic-biannual", name:"Basic", cycle:"Every 6 Months", price:80000,
      audience:"Small cooperatives and single-branch SACCOs just getting started",
      features:["Up to 200 members","Core registers: Members, Contributions, Loans, Repayments, Dividends","2 staff logins","Member communication (Email/SMS/WhatsApp)","Downloadable role-based user manual","Email support"] },
    { id:"professional-biannual", name:"Professional", cycle:"Every 6 Months", price:180000,
      audience:"Growing SACCOs and multi-officer teams that need full team access",
      features:["Unlimited members","Everything in Basic","Unlimited staff logins & role assignment","Priority support","Data export"] }
  ],
  yearly: [
    { id:"basic-yearly", name:"Basic", cycle:"Yearly", price:150000,
      audience:"Small cooperatives and single-branch SACCOs just getting started",
      features:["Up to 200 members","Core registers: Members, Contributions, Loans, Repayments, Dividends","2 staff logins","Member communication (Email/SMS/WhatsApp)","Downloadable role-based user manual","Email support"] },
    { id:"professional-yearly", name:"Professional", cycle:"Yearly", price:320000,
      audience:"Growing SACCOs and multi-officer teams that need full team access",
      features:["Unlimited members","Everything in Basic","Unlimited staff logins & role assignment","Priority support","Data export","2 months free vs. monthly billing"] }
  ]
};
/* Overwrites PLANS in place with the live catalogue from the Plans sheet
   (Super Admin > Plans & Pricing). Safe to call with no session — keeps the
   fallback catalogue above if offline or the request fails. */
async function loadPlans(){
  try{
    const result = await apiCall("listPlans", {});
    if(result.plans){
      Object.keys(PLANS).forEach(k=> delete PLANS[k]);
      Object.assign(PLANS, result.plans);
    }
  }catch(e){ /* keep fallback catalogue */ }
}
function findPlan(id){
  for(const cycle of Object.keys(PLANS)){ const hit = PLANS[cycle].find(p=> p.id===id); if(hit) return hit; }
  return null;
}


/* Roles (besides Admin) allowed to send member communications.
   Keep this in sync with MESSAGING_ROLES in Code.gs — this copy only
   controls whether the Communicate tab is shown; the server enforces it. */
const MESSAGING_ROLES = ["SACCO Manager / General Manager", "Membership Officer", "Accountant / Finance Officer"];
function canMessageClient(session){
  return session.isAdmin || (session.roles || []).some(r=> MESSAGING_ROLES.includes(r));
}

/* Adds a show/hide (eye) icon to every password field on the page. Self-
   contained: injects its own CSS once, so no page's <style> block needs to
   change. Safe to call more than once (skips fields already toggled) and
   safe on pages with no password fields at all. Relies on Bootstrap Icons
   (bi-eye / bi-eye-slash) — every page with a password field already loads
   that stylesheet. */
function initPasswordToggles(){
  if(!document.getElementById("pwToggleStyles")){
    const style = document.createElement("style");
    style.id = "pwToggleStyles";
    style.textContent = `
      .pw-toggle-wrap{position:relative;display:block;width:100%;}
      .pw-toggle-wrap input{padding-right:38px !important;}
      .pw-toggle-btn{position:absolute;top:50%;right:8px;transform:translateY(-50%);
        background:none;border:none;padding:4px;margin:0;line-height:1;color:#6B7C93;
        cursor:pointer;display:flex;align-items:center;}
      .pw-toggle-btn:hover{color:#1B6CA8;}
      .pw-toggle-btn:focus{outline:none;}`;
    document.head.appendChild(style);
  }
  document.querySelectorAll('input[type="password"]').forEach(input=>{
    if(input.dataset.pwToggled) return;
    input.dataset.pwToggled = "1";
    const wrap = document.createElement("span");
    wrap.className = "pw-toggle-wrap";
    input.parentNode.insertBefore(wrap, input);
    wrap.appendChild(input);
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "pw-toggle-btn";
    btn.tabIndex = -1; // keep tab order on the password field itself, not the icon
    btn.setAttribute("aria-label", "Show password");
    btn.innerHTML = '<i class="bi bi-eye"></i>';
    btn.addEventListener("click", ()=>{
      const showing = input.type === "password";
      input.type = showing ? "text" : "password";
      btn.innerHTML = showing ? '<i class="bi bi-eye-slash"></i>' : '<i class="bi bi-eye"></i>';
      btn.setAttribute("aria-label", showing ? "Hide password" : "Show password");
    });
    wrap.appendChild(btn);
  });
}
if(document.readyState === "loading"){
  document.addEventListener("DOMContentLoaded", initPasswordToggles);
} else {
  initPasswordToggles();
}

