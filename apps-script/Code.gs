/****************************************************************************
 * OKV SACCO — "Subscribe" (Online Multi-Tenant) Backend
 * ---------------------------------------------------------------------
 * Deploy: Extensions > Apps Script in a Google Sheet, paste this file,
 * run setupSheets() once, then Deploy > New deployment > Web app
 *   Execute as:  Me
 *   Who has access: Anyone
 * Copy the resulting /exec URL into config.js as API_URL.
 *
 * ARCHITECTURE — per-tenant spreadsheets (added):
 *   The spreadsheet this script is bound to is the MASTER/CONTROL
 *   spreadsheet (SS below). It holds only platform-wide, bounded data —
 *   never per-org business records:
 *     Tenants       — registry: orgId, orgName, spreadsheetId/Url of that
 *                     org's own data spreadsheet, createdAt, status
 *     Users         — one row per login (Admin or User), scoped by orgId.
 *                     Doubles as the platform's account-credentials index
 *                     (password hash, session/reset tokens) so login can
 *                     resolve a user before knowing which org they're in.
 *                     Also carries each org's plan/subscription fields.
 *     CellCapacity  — latest cell-usage snapshot per tenant (written by
 *                     checkAllTenantCellCapacity(), read by the Super
 *                     Admin dashboard's Capacity tab only)
 *     Payments, Messages, AdminMessages, Settings, Plans — unchanged;
 *     these are operational/platform logs (bounded by org & staff count,
 *     not by business-transaction volume), not the "Data" register.
 *
 *   Every org's actual business records (Members/Contributions/Loans/
 *   Repayments/Dividends — the "Data" sheet, the one that grows without
 *   bound as an org uses the system and is what can approach Google
 *   Sheets' 10,000,000-cell-per-file limit) live in a SEPARATE spreadsheet
 *   created automatically for that org at signup. Every function that
 *   touches the Data sheet resolves that spreadsheet dynamically via
 *   getOrgSpreadsheet_(orgId) — nothing hardcodes or assumes "the active
 *   spreadsheet" for business data. See getOrgSpreadsheet_(),
 *   createTenantSpreadsheet_(), and migrateExistingDataToPerTenant_()
 *   below.
 ****************************************************************************/

const SS = SpreadsheetApp.getActiveSpreadsheet(); // the MASTER/CONTROL spreadsheet
const SESSION_HOURS = 12;
const RESET_TOKEN_MINUTES = 60;

/* ================= OKV Technology Consults' own outbound email identity =================
   Applies ONLY to emails the PLATFORM sends on its own behalf — welcome
   emails, password resets, payment/subscription notices, trial reminders,
   and Super Admin announcements to org Admins. It must NOT be used for
   emails an organization sends to its own members or staff (handleSendMessage_,
   handleSendStatements_, handleSendTeamMessage_) — those carry that
   organization's own identity, not OKV's. */
const OKV_SENDER_NAME = "OKV SACCO Subscribe - OKV Technology Consults";
const OKV_SENDER_EMAIL = "technologyokv@gmail.com";
function sendOkvEmail_(to, subject, body, htmlBody){
  const options = { name: OKV_SENDER_NAME, replyTo: OKV_SENDER_EMAIL };
  if(htmlBody) options.htmlBody = htmlBody;
  MailApp.sendEmail(to, subject, body, options);
}

const USERS_HEADERS = ["orgId","orgName","userId","username","email","fullName","phone","passwordHash",
  "roles","isAdmin","status","subscriptionStatus","subscriptionExpiry","plan","billingCycle","onTrial",
  "resetToken","resetTokenExpiry","sessionToken","sessionExpiry","createdAt","isSuperAdmin",
  "lastReminderSent","reminderCycleExpiry","remindersSentThisCycle"];
const DATA_HEADERS = ["orgId","recordType","recordId","createdBy","payload","updatedAt","deleted"];
const ORG_PROFILE_HEADERS = ["key","value"];
const PAYMENTS_HEADERS = ["paymentId","orgId","userId","orgName","fullName","email","phone",
  "planId","planName","amount","proofUrl","status","notified","submittedAt","reviewedAt"];
const MESSAGES_HEADERS = ["msgId","orgId","userId","subject","message","channels",
  "recipientCount","sent","failed","sentAt"];
const ADMIN_MESSAGES_HEADERS = ["msgId","sentBy","subject","message","channels",
  "recipientCount","sent","failed","sentAt"];
const SETTINGS_HEADERS = ["key","value"];
const PLANS_HEADERS = ["planId","name","cycle","cycleKey","price","audience","features","active","sortOrder"];
const TENANTS_HEADERS = ["orgId","orgName","spreadsheetId","spreadsheetUrl","createdAt","status"];
const CELL_CAPACITY_HEADERS = ["orgId","orgName","spreadsheetId","totalCellsUsed","percentUsed",
  "rowsAdded30d","rowsAdded90d","growthRatePerDay","estDaysTo90Percent","statusFlag",
  "recommendedAction","lastChecked","historySnapshotJson"];

const CELL_CAPACITY_LIMIT = 10000000; // Google Sheets' hard per-file cell cap
const CELL_CAPACITY_MONITOR_THRESHOLD = 0.70;  // >= 70% used  -> "Monitor"
const CELL_CAPACITY_ACTION_THRESHOLD  = 0.85;  // >= 85% used  -> "Action Needed"
const CELL_CAPACITY_PROJECTION_TARGET = 0.90;  // project days remaining until this level

const REMINDER_THRESHOLD_DAYS = 14; // only consider sending once expiry is this close
const REMINDER_SEND_DAYS_OF_MONTH = [1, 15]; // reminders only ever go out on these two calendar days

/* Default values seeded into the Settings sheet the first time setupSheets()
   runs. After that, edit them from the Super Admin dashboard's Settings tab
   instead — these constants are only a one-time starting point. */
const DEFAULT_SETTINGS = {
  contactWhatsapp: "+2348104141138", contactPhone: "+2348104141138",
  contactEmail: "technologyokv@gmail.com", contactWebsite: "www.okvtechnology.com",
  bankAccountNumber: "8104141138", bankName: "OPay MFB", bankAccountName: "Olasile Kehinde Victor",
  notifyEmail: "technologyokv@gmail.com", manualPaymentEnabled: "true",

  brandPrimaryColor: "#22C7A9", brandSecondaryColor: "#0D1B2A", brandLogoUrl: "",
  systemName: "OKV SACCO", tagline: "Savings & Credit Cooperative Society",

  smsProvider: "termii", smsApiKey: "", smsSenderId: "OKVSACCO",
  whatsappProvider: "termii", whatsappApiKey: "", whatsappSenderId: "OKVSACCO",

  paystackEnabled: "false", paystackPublicKey: "", paystackSecretKey: "",
  remitaEnabled: "false", remitaPublicKey: "", remitaSecretKey: "",
  flutterwaveEnabled: "false", flutterwavePublicKey: "", flutterwaveSecretKey: "",

  trialDays: "7"
};

function setupSheets(){
  ensureSheet_("Users", USERS_HEADERS);
  ensureSheet_("Payments", PAYMENTS_HEADERS);
  ensureSheet_("Messages", MESSAGES_HEADERS);
  ensureSheet_("AdminMessages", ADMIN_MESSAGES_HEADERS);
  ensureSheet_("Settings", SETTINGS_HEADERS);
  ensureSheet_("Plans", PLANS_HEADERS);
  ensureSheet_("Tenants", TENANTS_HEADERS);
  ensureSheet_("CellCapacity", CELL_CAPACITY_HEADERS);
  seedDefaultSettings_();
  seedDefaultPlans_();
}
function seedDefaultPlans_(){
  if(readRows_("Plans").length > 0) return; // already has data — never overwrite edits
  const sh = sheet_("Plans");
  const defaults = [
    ["basic-monthly","Basic","Monthly","monthly",15000,"Small cooperatives and single-branch SACCOs just getting started",
      JSON.stringify(["Up to 200 members","Core registers: Members, Contributions, Loans, Repayments, Dividends","2 staff logins","Member communication (Email/SMS/WhatsApp)","Downloadable role-based user manual","Email support"]), true, 1],
    ["professional-monthly","Professional","Monthly","monthly",35000,"Growing SACCOs and multi-officer teams that need full team access",
      JSON.stringify(["Unlimited members","Everything in Basic","Unlimited staff logins & role assignment","Priority support","Data export"]), true, 2],
    ["basic-biannual","Basic","Every 6 Months","biannual",80000,"Small cooperatives and single-branch SACCOs just getting started",
      JSON.stringify(["Up to 200 members","Core registers: Members, Contributions, Loans, Repayments, Dividends","2 staff logins","Member communication (Email/SMS/WhatsApp)","Downloadable role-based user manual","Email support"]), true, 1],
    ["professional-biannual","Professional","Every 6 Months","biannual",180000,"Growing SACCOs and multi-officer teams that need full team access",
      JSON.stringify(["Unlimited members","Everything in Basic","Unlimited staff logins & role assignment","Priority support","Data export"]), true, 2],
    ["basic-yearly","Basic","Yearly","yearly",150000,"Small cooperatives and single-branch SACCOs just getting started",
      JSON.stringify(["Up to 200 members","Core registers: Members, Contributions, Loans, Repayments, Dividends","2 staff logins","Member communication (Email/SMS/WhatsApp)","Downloadable role-based user manual","Email support"]), true, 1],
    ["professional-yearly","Professional","Yearly","yearly",320000,"Growing SACCOs and multi-officer teams that need full team access",
      JSON.stringify(["Unlimited members","Everything in Basic","Unlimited staff logins & role assignment","Priority support","Data export","2 months free vs. monthly billing"]), true, 2]
  ];
  defaults.forEach(row=> sh.appendRow(row));
}
function seedDefaultSettings_(){
  const sh = sheet_("Settings");
  const existing = readRows_("Settings");
  const existingKeys = {}; existing.forEach(r=> existingKeys[r.key] = true);
  Object.keys(DEFAULT_SETTINGS).forEach(key=>{
    if(!existingKeys[key]) sh.appendRow([key, DEFAULT_SETTINGS[key]]);
  });
}
function getSetting_(key){
  const row = readRows_("Settings").find(r=> r.key === key);
  return row ? row.value : (DEFAULT_SETTINGS[key] || "");
}
function ensureSheet_(name, headers, ss){
  const targetSS = ss || SS;
  let sh = targetSS.getSheetByName(name);
  if(!sh) sh = targetSS.insertSheet(name);
  if(sh.getLastRow() === 0){ sh.appendRow(headers); sh.setFrozenRows(1); return; }
  // Existing sheet: safely add any headers introduced since it was created,
  // without touching existing rows/columns. Re-run setupSheets() any time
  // this file gains new columns.
  const existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const missing = headers.filter(h=> existing.indexOf(h) === -1);
  if(missing.length) sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
}

/* Run this ONCE from the Apps Script editor (select "setupSuperAdmin" in the
   function dropdown, click Run) to create your own Super Admin login —
   change the three values below first. Delete/blank the password line
   afterwards if you like; the hash is what's actually stored. */
function setupSuperAdmin(){
  const username = "technologyokv@gmail.com";
  const email = "technologyokv@gmail.com";
  const plainPassword = "OKVSACCO557"; // default password — change this after your first sign-in

  const exists = findUserRow_(u=> String(u.username).toLowerCase() === username.toLowerCase());
  if(exists){ Logger.log("A user with that username already exists."); return; }
  sheet_("Users").appendRow(["PLATFORM", "OKV Technology Consults", "USR-" + Utilities.getUuid().slice(0,8).toUpperCase(),
    username, email, "Super Admin", "", hashPassword_(plainPassword), JSON.stringify([]), false, "Active",
    "active", "", "", "", false, "", "", "", "", new Date().toISOString(), true, "", "", 0]);
  Logger.log("Super Admin created — username: " + username);
}
function hashPassword_(text){
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return raw.map(b=> ((b < 0 ? b + 256 : b).toString(16)).padStart(2, "0")).join("");
}

/* ================= entry point ================= */
function doPost(e){
  let body;
  try{ body = JSON.parse(e.postData.contents); }
  catch(err){ return jsonOut_({ok:false, error:"Invalid request body."}); }

  const action = body.action;
  try{
    switch(action){
      case "signup":            return jsonOut_(handleSignup_(body));
      case "login":              return jsonOut_(handleLogin_(body));
      case "resumeSession":      return jsonOut_(handleResumeSession_(body));
      case "forgotPassword":     return jsonOut_(handleForgotPassword_(body));
      case "resetPassword":      return jsonOut_(handleResetPassword_(body));
      case "changePassword":     return jsonOut_(handleChangePassword_(body));
      case "listUsers":          return jsonOut_(handleListUsers_(body));
      case "createUser":         return jsonOut_(handleCreateUser_(body));
      case "updateUser":         return jsonOut_(handleUpdateUser_(body));
      case "deleteUser":         return jsonOut_(handleDeleteUser_(body));
      case "setUserPassword":    return jsonOut_(handleSetUserPassword_(body));
      case "setUserStatus":      return jsonOut_(handleSetUserStatus_(body));
      case "syncPush":           return jsonOut_(handleSyncPush_(body));
      case "syncPull":           return jsonOut_(handleSyncPull_(body));
      case "submitPayment":      return jsonOut_(handleSubmitPayment_(body));
      case "listPayments":       return jsonOut_(handleListPayments_(body));
      case "sendMessage":        return jsonOut_(handleSendMessage_(body));
      case "listMessages":       return jsonOut_(handleListMessages_(body));
      case "sendStatements":     return jsonOut_(handleSendStatements_(body));
      case "sendTeamMessage":    return jsonOut_(handleSendTeamMessage_(body));
      case "getOrgProfile":      return jsonOut_(handleGetOrgProfile_(body));
      case "updateOrgProfile":   return jsonOut_(handleUpdateOrgProfile_(body));
      case "listOrganizations":  return jsonOut_(handleListOrganizations_(body));
      case "updateOrganization": return jsonOut_(handleUpdateOrganization_(body));
      case "listAllPayments":    return jsonOut_(handleListAllPayments_(body));
      case "reviewPayment":      return jsonOut_(handleReviewPayment_(body));
      case "superSendMessage":   return jsonOut_(handleSuperSendMessage_(body));
      case "listAdminMessages":  return jsonOut_(handleListAdminMessages_(body));
      case "getSettings":        return jsonOut_(handleGetSettings_(body));
      case "getSettingsFull":    return jsonOut_(handleGetSettingsFull_(body));
      case "updateSettings":     return jsonOut_(handleUpdateSettings_(body));
      case "uploadLogo":         return jsonOut_(handleUploadLogo_(body));
      case "listPlans":          return jsonOut_(handleListPlans_(body));
      case "listAllPlans":       return jsonOut_(handleListAllPlans_(body));
      case "savePlan":           return jsonOut_(handleSavePlan_(body));
      case "deletePlan":         return jsonOut_(handleDeletePlan_(body));
      case "getCapacityReport":  return jsonOut_(handleGetCapacityReport_(body));
      default: return jsonOut_({ok:false, error:"Unknown action: " + action});
    }
  }catch(err){
    return jsonOut_({ok:false, error: err.message});
  }
}
function doGet(e){
  return jsonOut_({ok:true, message:"OKV SACCO Subscribe API is running. Use POST."});
}
function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= sheet helpers =================
   Every helper here optionally takes a target spreadsheet (ss). Omit it and
   these operate on the MASTER/CONTROL spreadsheet (SS), exactly as before —
   only the call sites that touch per-tenant "Data" pass one explicitly, via
   getOrgSpreadsheet_(orgId). Nothing else changes behavior. */
function sheet_(name, ss){ return (ss || SS).getSheetByName(name); }
function readRows_(sheetName, ss){
  const sh = sheet_(sheetName, ss);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return values.map((row, idx)=>{
    const obj = {}; headers.forEach((h,i)=> obj[h]=row[i]);
    obj.__row = idx + 2; // +1 for header row, +1 because sheet rows are 1-indexed
    return obj;
  });
}
/* ================= per-tenant spreadsheet resolution =================
   Central helper: every function that needs an org's business-data
   spreadsheet calls this instead of assuming "the active spreadsheet".
   Looks up the org in the master Tenants registry and opens it by ID. */
function getOrgSpreadsheet_(orgId){
  const tenant = readRows_("Tenants").find(t=> t.orgId === orgId);
  if(!tenant || !tenant.spreadsheetId){
    throw new Error("No data spreadsheet is registered for this organization. Contact support.");
  }
  return SpreadsheetApp.openById(tenant.spreadsheetId);
}
/* Creates a brand-new per-tenant spreadsheet (Data sheet only), registers
   it in the master Tenants sheet, and returns it. Called from signup and
   from the one-time migration script. */
function createTenantSpreadsheet_(orgId, orgName){
  const ss = SpreadsheetApp.create((orgName || orgId) + " — OKV SACCO Data");
  ensureSheet_("Data", DATA_HEADERS, ss);
  ensureSheet_("OrgProfile", ORG_PROFILE_HEADERS, ss);
  const defaultSheet = ss.getSheetByName("Sheet1");
  if(defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  sheet_("Tenants").appendRow([orgId, orgName || "", ss.getId(), ss.getUrl(), new Date().toISOString(), "Active"]);
  return ss;
}
/* Reads a per-tenant OrgProfile key/value sheet into a plain object. Falls
   back to {} for tenants created before OrgProfile existed. */
function getOrgProfile_(orgId){
  const ss = getOrgSpreadsheet_(orgId);
  if(!ss.getSheetByName("OrgProfile")) ensureSheet_("OrgProfile", ORG_PROFILE_HEADERS, ss);
  const rows = readRows_("OrgProfile", ss);
  const profile = {};
  rows.forEach(r=> profile[r.key] = r.value);
  return profile;
}
function setOrgProfileValue_(orgId, key, value){
  const ss = getOrgSpreadsheet_(orgId);
  if(!ss.getSheetByName("OrgProfile")) ensureSheet_("OrgProfile", ORG_PROFILE_HEADERS, ss);
  const sh = sheet_("OrgProfile", ss);
  const rows = readRows_("OrgProfile", ss);
  const existing = rows.find(r=> r.key === key);
  if(existing) sh.getRange(existing.__row, 2).setValue(value);
  else sh.appendRow([key, value]);
}
function handleGetOrgProfile_(body){
  const caller = verifySession_(body);
  return { ok:true, profile: getOrgProfile_(caller.orgId) };
}
function handleUpdateOrgProfile_(body){
  const caller = verifySession_(body);
  if(!caller.isAdmin) throw new Error("Only your organization's Admin can update organization details.");
  const fields = ["displayName","logoUrl","address","phone","email","currencySymbol"];
  fields.forEach(f=>{
    if(body.profile && Object.prototype.hasOwnProperty.call(body.profile, f)){
      setOrgProfileValue_(caller.orgId, f, String(body.profile[f] || ""));
    }
  });
  return { ok:true, profile: getOrgProfile_(caller.orgId) };
}
function findUserRow_(predicate){
  return readRows_("Users").find(predicate) || null;
}
function updateUserRow_(rowIndex, updates){
  const sh = sheet_("Users");
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  Object.keys(updates).forEach(k=>{
    const col = headers.indexOf(k) + 1;
    if(col > 0) sh.getRange(rowIndex, col).setValue(updates[k]);
  });
}
function isTrue_(v){ return v===true || v==="TRUE" || v===1; }
function newToken_(){ return Utilities.getUuid(); }

/* ================= auth ================= */
function verifySession_(body){
  const user = findUserRow_(u=> u.userId===body.userId && u.orgId===body.orgId && String(u.sessionToken)===String(body.sessionToken) && u.sessionToken);
  if(!user) throw new Error("Session invalid. Please log in again.");
  if(new Date(user.sessionExpiry) < new Date()) throw new Error("Session expired. Please log in again.");
  if(user.status !== "Active") throw new Error("This account has been suspended. Contact your Admin.");
  return user;
}
function requireAdmin_(caller){
  if(!isTrue_(caller.isAdmin)) throw new Error("Only an Admin can do this.");
}
function sessionPayload_(user, token, expiry){
  return {
    ok:true, orgId:user.orgId, orgName:user.orgName||"", userId:user.userId, username:user.username,
    email:user.email, fullName:user.fullName||"", phone:user.phone||"",
    roles: JSON.parse(user.roles || "[]"), isAdmin: isTrue_(user.isAdmin), isSuperAdmin: isTrue_(user.isSuperAdmin),
    subscriptionStatus: user.subscriptionStatus || "", subscriptionExpiry: user.subscriptionExpiry || "",
    plan: user.plan || "", billingCycle: user.billingCycle || "", onTrial: isTrue_(user.onTrial),
    sessionToken: token, sessionExpiry: expiry.toISOString()
  };
}

function handleSignup_(body){
  const { orgName, fullName, email, phone, passwordHash, plan, billingCycle } = body;
  if(!email || !passwordHash) throw new Error("Email and password are required.");
  const username = String(email).toLowerCase();
  const exists = findUserRow_(u=> String(u.username).toLowerCase() === username);
  if(exists) throw new Error("An account with that email already exists.");
  const orgId = "ORG-" + Utilities.getUuid().slice(0,8).toUpperCase();
  const userId = "USR-" + Utilities.getUuid().slice(0,8).toUpperCase();
  const now = new Date();
  const trialDays = Number(getSetting_("trialDays")) || 7;
  const trialExpiry = new Date(now.getTime() + trialDays*24*60*60*1000);
  try{
    createTenantSpreadsheet_(orgId, orgName);
  }catch(err){
    throw new Error("Could not set up your organization's data spreadsheet: " + err.message);
  }
  sheet_("Users").appendRow([orgId, orgName||"", userId, username, email, fullName||"", phone||"", passwordHash,
    JSON.stringify(["SACCO Super Admin"]), true, "Active", "active", trialExpiry.toISOString(),
    plan||"", billingCycle||"", true, "", "", "", "", now.toISOString(), false, "", "", 0]);
  try{
    sendOkvEmail_(email, "Welcome to OKV SACCO — your " + trialDays + "-day free trial has started",
      "Hello " + (fullName || username) + ",\n\n" +
      "Your organization \"" + (orgName || "") + "\" is ready.\n" +
      "Plan selected: " + (plan || "Not specified") + " (" + (billingCycle || "N/A") + ")\n" +
      "Your free trial is active for 7 days, expiring " + trialExpiry.toDateString() + ".\n\n" +
      "Next step: install the app and sign in with this email and the password you chose.\n\n" +
      "— OKV SACCO Management System");
  }catch(e){ /* email is best-effort — don't fail signup if MailApp quota/permissions block it */ }
  return { ok:true, orgId, userId };
}

function handleLogin_(body){
  const { username, passwordHash } = body;
  const user = findUserRow_(u=> String(u.username).toLowerCase() === String(username||"").toLowerCase());
  if(!user || String(user.passwordHash) !== String(passwordHash)) throw new Error("Incorrect username or password.");
  if(user.status !== "Active") throw new Error("This account has been suspended. Contact your Admin.");
  if(!isTrue_(user.isSuperAdmin) && user.subscriptionStatus !== "active") throw new Error("Subscription inactive. Contact your Admin.");
  const token = newToken_();
  const expiry = new Date(Date.now() + SESSION_HOURS*60*60*1000);
  updateUserRow_(user.__row, { sessionToken: token, sessionExpiry: expiry.toISOString() });
  return sessionPayload_(user, token, expiry);
}

function handleResumeSession_(body){
  const user = verifySession_(body);
  const expiry = new Date(Date.now() + SESSION_HOURS*60*60*1000);
  updateUserRow_(user.__row, { sessionExpiry: expiry.toISOString() });
  return sessionPayload_(user, body.sessionToken, expiry);
}

function handleForgotPassword_(body){
  const q = String(body.username || "").toLowerCase();
  // Self-service reset is for Admins only, by design — a User has an Admin
  // above them who resets their password directly from the Users tab instead.
  const user = findUserRow_(u=> (String(u.username).toLowerCase()===q || String(u.email).toLowerCase()===q) && isTrue_(u.isAdmin));
  if(!user) return { ok:true }; // don't reveal whether an account exists (or that it's a non-Admin)
  const token = newToken_();
  const expiry = new Date(Date.now() + RESET_TOKEN_MINUTES*60*1000);
  updateUserRow_(user.__row, { resetToken: token, resetTokenExpiry: expiry.toISOString() });
  const link = body.resetBaseUrl + "?token=" + token;
  sendOkvEmail_((user.email || user.username),
    "Reset your OKV SACCO password",
    "Hello " + user.username + ",\n\n" +
    "A password reset was requested for your account.\n" +
    "Reset your password using the link below (valid for 1 hour):\n\n" + link +
    "\n\nIf you did not request this, you can safely ignore this email.\n\n— OKV SACCO Management System");
  return { ok:true };
}

function handleResetPassword_(body){
  const { token, newPasswordHash } = body;
  if(!token) throw new Error("Missing reset token.");
  const user = findUserRow_(u=> u.resetToken && String(u.resetToken) === String(token));
  if(!user) throw new Error("This reset link is invalid or has already been used.");
  if(new Date(user.resetTokenExpiry) < new Date()) throw new Error("This reset link has expired. Please request a new one.");
  updateUserRow_(user.__row, { passwordHash:newPasswordHash, resetToken:"", resetTokenExpiry:"" });
  return { ok:true };
}

function handleChangePassword_(body){
  const user = verifySession_(body);
  if(String(user.passwordHash) !== String(body.currentPasswordHash)) throw new Error("Current password is incorrect.");
  updateUserRow_(user.__row, { passwordHash: body.newPasswordHash });
  return { ok:true };
}

/* ================= admin: user management (scoped to caller's orgId) ================= */
function publicUser_(u){
  return { userId:u.userId, username:u.username, email:u.email, phone:u.phone, fullName:u.fullName, roles: JSON.parse(u.roles || "[]"),
    isAdmin: isTrue_(u.isAdmin), status:u.status, subscriptionStatus:u.subscriptionStatus,
    subscriptionExpiry:u.subscriptionExpiry, createdAt:u.createdAt };
}
function handleListUsers_(body){
  const caller = verifySession_(body);
  const rows = readRows_("Users").filter(u=> u.orgId === caller.orgId);
  return { ok:true, users: rows.map(publicUser_) };
}
function handleCreateUser_(body){
  const caller = verifySession_(body); requireAdmin_(caller);
  if(!body.username || !body.passwordHash) throw new Error("Username and password are required.");
  const exists = findUserRow_(u=> String(u.username).toLowerCase() === String(body.username).toLowerCase());
  if(exists) throw new Error("That username is already taken.");
  const userId = "USR-" + Utilities.getUuid().slice(0,8).toUpperCase();
  sheet_("Users").appendRow([caller.orgId, "", userId, body.username, body.email||"", body.fullName||"", body.phone||"",
    body.passwordHash, JSON.stringify(body.roles||[]), false, "Active", "active", "", "", "", false,
    "", "", "", "", new Date().toISOString(), false, "", "", 0]);
  return { ok:true, userId };
}
function handleUpdateUser_(body){
  const caller = verifySession_(body); requireAdmin_(caller);
  const target = findUserRow_(u=> u.userId===body.userId && u.orgId===caller.orgId);
  if(!target) throw new Error("User not found in your organization.");
  const updates = {};
  if(body.username !== undefined && body.username !== target.username){
    const clash = findUserRow_(u=> String(u.username).toLowerCase() === String(body.username).toLowerCase() && u.userId !== target.userId);
    if(clash) throw new Error("That username is already taken.");
    updates.username = body.username;
  }
  if(body.email !== undefined) updates.email = body.email;
  if(body.fullName !== undefined) updates.fullName = body.fullName;
  if(body.phone !== undefined) updates.phone = body.phone;
  if(body.roles !== undefined) updates.roles = JSON.stringify(body.roles);
  updateUserRow_(target.__row, updates);
  return { ok:true };
}
function handleDeleteUser_(body){
  const caller = verifySession_(body); requireAdmin_(caller);
  const target = findUserRow_(u=> u.userId===body.userId && u.orgId===caller.orgId);
  if(!target) throw new Error("User not found in your organization.");
  if(isTrue_(target.isAdmin)) throw new Error("The Admin account cannot be deleted.");
  sheet_("Users").deleteRow(target.__row);
  return { ok:true };
}
function handleSetUserPassword_(body){
  const caller = verifySession_(body); requireAdmin_(caller);
  const target = findUserRow_(u=> u.userId===body.userId && u.orgId===caller.orgId);
  if(!target) throw new Error("User not found in your organization.");
  updateUserRow_(target.__row, { passwordHash: body.newPasswordHash });
  return { ok:true };
}
function handleSetUserStatus_(body){
  const caller = verifySession_(body); requireAdmin_(caller);
  const target = findUserRow_(u=> u.userId===body.userId && u.orgId===caller.orgId);
  if(!target) throw new Error("User not found in your organization.");
  if(isTrue_(target.isAdmin)) throw new Error("The Admin account cannot be suspended.");
  updateUserRow_(target.__row, { status: body.status }); // "Active" | "Suspended"
  return { ok:true };
}

/* ================= data sync (generic — covers every register) ================= */
function handleSyncPush_(body){
  const caller = verifySession_(body);
  const orgSS = getOrgSpreadsheet_(caller.orgId);
  const records = body.records || [];
  const sh = sheet_("Data", orgSS);
  const existing = readRows_("Data", orgSS);
  const index = {}; existing.forEach(r=> index[r.recordType + "::" + r.recordId] = r);
  let applied = 0;
  records.forEach(rec=>{
    const key = rec.recordType + "::" + rec.recordId;
    const current = index[key];
    if(!current || new Date(rec.updatedAt) > new Date(current.updatedAt)){
      const rowValues = [caller.orgId, rec.recordType, rec.recordId, rec.createdBy||"",
        JSON.stringify(rec.payload||{}), rec.updatedAt, rec.deleted ? true : false];
      if(current) sh.getRange(current.__row, 1, 1, DATA_HEADERS.length).setValues([rowValues]);
      else sh.appendRow(rowValues);
      applied++;
    }
  });
  return { ok:true, applied, serverTime: new Date().toISOString() };
}
function handleSyncPull_(body){
  const caller = verifySession_(body);
  const since = body.since ? new Date(body.since) : new Date(0);
  const rows = readRows_("Data", getOrgSpreadsheet_(caller.orgId)).filter(r=> new Date(r.updatedAt) > since);
  const records = rows.map(r=> ({
    recordType: r.recordType, recordId: r.recordId, createdBy: r.createdBy,
    payload: JSON.parse(r.payload || "{}"), updatedAt: r.updatedAt, deleted: isTrue_(r.deleted)
  }));
  return { ok:true, records, serverTime: new Date().toISOString() };
}

/* ================= subscription payment submission ================= */
function updateSheetCell_(sheetName, rowIndex, colName, value){
  const sh = sheet_(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const col = headers.indexOf(colName) + 1;
  if(col > 0) sh.getRange(rowIndex, col).setValue(value);
}
function getOrCreateFolder_(name){
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function handleSubmitPayment_(body){
  const caller = verifySession_(body); requireAdmin_(caller);
  const { orgName, fullName, email, phone, planId, planName, amount, proofBase64, proofFileName } = body;
  let proofUrl = "";
  if(proofBase64){
    try{
      const folder = getOrCreateFolder_("OKV SACCO Payment Proofs");
      const commaIdx = proofBase64.indexOf(",");
      const raw = commaIdx > -1 ? proofBase64.slice(commaIdx+1) : proofBase64;
      const blob = Utilities.newBlob(Utilities.base64Decode(raw), "image/png", proofFileName || (caller.orgId + "-proof.png"));
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      proofUrl = file.getUrl();
    }catch(err){ /* proof upload failing shouldn't block the submission */ }
  }
  const paymentId = "PAY-" + Utilities.getUuid().slice(0,8).toUpperCase();
  sheet_("Payments").appendRow([paymentId, caller.orgId, caller.userId, orgName||caller.orgName||"",
    fullName||caller.fullName||caller.username, email||caller.email||"", phone||caller.phone||"",
    planId||"", planName||"", amount||"", proofUrl, "Pending", false, new Date().toISOString(), ""]);
  sendOkvEmail_(getSetting_("notifyEmail"), "New subscription payment submitted — " + (orgName || caller.orgId),
    "Organization: " + (orgName || caller.orgId) + "\n" +
    "Org ID: " + caller.orgId + "\n" +
    "Contact: " + (fullName || caller.username) + " (" + (email || caller.email) + ", " + (phone||"") + ")\n" +
    "Plan: " + planName + "\n" +
    "Amount: " + amount + "\n" +
    "Payment proof: " + (proofUrl || "not provided") + "\n\n" +
    "Review in the Payments sheet and set status to Confirmed or Rejected — the requester is emailed " +
    "automatically the next time checkPaymentUpdates() runs.");
  return { ok:true, paymentId };
}
function handleListPayments_(body){
  const caller = verifySession_(body); requireAdmin_(caller);
  const rows = readRows_("Payments").filter(p=> p.orgId === caller.orgId)
    .sort((a,b)=> String(b.submittedAt).localeCompare(String(a.submittedAt)));
  return { ok:true, payments: rows.map(p=> ({ paymentId:p.paymentId, planName:p.planName, amount:p.amount,
    status:p.status, submittedAt:p.submittedAt, reviewedAt:p.reviewedAt })) };
}
function planDurationDays_(planId){
  if(!planId) return 30;
  if(planId.indexOf("yearly") > -1) return 365;
  if(planId.indexOf("biannual") > -1) return 182;
  return 30;
}
/* Set up a time-driven trigger for this function (every 10-15 minutes) so
   payment confirmations/rejections reach the user well within an hour:
   Apps Script editor > Triggers (clock icon) > Add Trigger >
   Function: checkPaymentUpdates, Event source: Time-driven, Every 10 minutes. */
function checkPaymentUpdates(){
  const rows = readRows_("Payments");
  rows.forEach(r=>{
    if((r.status === "Confirmed" || r.status === "Rejected") && !isTrue_(r.notified)){
      const user = findUserRow_(u=> u.userId === r.userId && u.orgId === r.orgId);
      if(user){
        if(r.status === "Confirmed"){
          const days = planDurationDays_(r.planId);
          const expiry = new Date(Date.now() + days*24*60*60*1000);
          updateUserRow_(user.__row, { subscriptionStatus:"active", subscriptionExpiry: expiry.toISOString(),
            plan: r.planId, onTrial:false });
        }
        sendOkvEmail_(r.email || user.email || user.username,
          "Your OKV SACCO subscription payment — " + r.status,
          r.status === "Confirmed"
            ? "Good news! Your payment has been confirmed and your subscription is now active.\nSign in to your dashboard to see your new plan and renewal date.\n\n— OKV Technology Consults"
            : "We could not confirm your payment. Please reach out to us on WhatsApp or email so we can help resolve this.\n\n— OKV Technology Consults");
      }
      updateSheetCell_("Payments", r.__row, "notified", true);
      updateSheetCell_("Payments", r.__row, "reviewedAt", new Date().toISOString());
    }
  });
}

/* ================= member communication (email / SMS / WhatsApp) ================= */
const MESSAGING_ROLES = ["SACCO Manager / General Manager", "Membership Officer", "Accountant / Finance Officer"];
function canMessage_(caller){
  if(isTrue_(caller.isAdmin)) return true;
  const roles = JSON.parse(caller.roles || "[]");
  return roles.some(r=> MESSAGING_ROLES.indexOf(r) > -1);
}
function sendViaTermii_(phone, message, channel, apiKey, senderId){
  if(!apiKey) return false;
  try{
    const url = channel === "whatsapp" ? "https://api.ng.termii.com/api/send/whatsapp" : "https://api.ng.termii.com/api/sms/send";
    const payload = { to: phone, from: senderId || "OKVSACCO", sms: message, type:"plain",
      channel: channel === "whatsapp" ? "whatsapp" : "generic", api_key: apiKey };
    const res = UrlFetchApp.fetch(url, { method:"post", contentType:"application/json",
      payload: JSON.stringify(payload), muteHttpExceptions:true });
    return res.getResponseCode() < 300;
  }catch(err){ return false; }
}
/* Routes to whichever provider is configured in Settings for this channel.
   Currently wired for Termii (covers both SMS and WhatsApp); add another
   provider's branch here if you switch — the calling code never changes. */
function sendViaChannel_(phone, message, channel){
  const provider = getSetting_(channel === "whatsapp" ? "whatsappProvider" : "smsProvider");
  const apiKey = getSetting_(channel === "whatsapp" ? "whatsappApiKey" : "smsApiKey");
  const senderId = getSetting_(channel === "whatsapp" ? "whatsappSenderId" : "smsSenderId");
  if(!apiKey) return false; // not configured yet — safely no-op, Email still works
  if(provider === "termii" || !provider) return sendViaTermii_(phone, message, channel, apiKey, senderId);
  return false; // unknown provider configured
}
function handleSendMessage_(body){
  const caller = verifySession_(body);
  if(!canMessage_(caller)) throw new Error("You don't have permission to send messages.");
  const { subject, message, memberIds, allMembers, channels } = body;
  if(!message) throw new Error("Message body is required.");
  if(!channels || !channels.length) throw new Error("Choose at least one channel (Email, SMS, or WhatsApp).");
  const memberRows = readRows_("Data", getOrgSpreadsheet_(caller.orgId)).filter(r=> r.recordType === "members" && !isTrue_(r.deleted));
  let targets = memberRows.map(r=> JSON.parse(r.payload || "{}"));
  if(!allMembers){
    const idSet = {}; (memberIds || []).forEach(id=> idSet[id] = true);
    targets = targets.filter(m=> idSet[m.recordId]);
  }
  if(!targets.length) throw new Error("No matching members to message.");
  let sent = 0, failed = 0;
  targets.forEach(m=>{
    const personalized = String(message).replace(/\{name\}/g, m.name || "Member").replace(/\{memberId\}/g, m.recordId || "");
    if(channels.indexOf("email") > -1){
      if(m.email){ try{ MailApp.sendEmail(m.email, subject || "Message from your SACCO", personalized); sent++; }catch(e){ failed++; } }
      else failed++;
    }
    if(channels.indexOf("sms") > -1){
      if(m.phone && sendViaChannel_(m.phone, personalized, "sms")) sent++; else failed++;
    }
    if(channels.indexOf("whatsapp") > -1){
      if(m.phone && sendViaChannel_(m.phone, personalized, "whatsapp")) sent++; else failed++;
    }
  });
  sheet_("Messages").appendRow([Utilities.getUuid().slice(0,8).toUpperCase(), caller.orgId, caller.userId,
    subject || "", message, channels.join(","), targets.length, sent, failed, new Date().toISOString()]);
  return { ok:true, recipients: targets.length, sent, failed };
}
function handleSendStatements_(body){
  const caller = verifySession_(body);
  if(!canMessage_(caller)) throw new Error("You don't have permission to send statements.");
  const statements = body.statements || [];
  if(!statements.length) throw new Error("No statements to send.");
  let sent = 0, failed = 0;
  statements.forEach(s=>{
    const channel = s.channel || "email";
    try{
      if(channel === "email"){
        if(!s.to) throw new Error("no email");
        MailApp.sendEmail({ to: s.to, subject: s.subject || "Your Statement", htmlBody: s.htmlBody, noReply: true });
      } else {
        if(!s.phone) throw new Error("no phone");
        if(!sendViaChannel_(s.phone, s.textBody || "", channel)) throw new Error("channel not configured");
      }
      sent++;
    }catch(e){ failed++; }
  });
  sheet_("Messages").appendRow([Utilities.getUuid().slice(0,8).toUpperCase(), caller.orgId, caller.userId,
    "Member Statement(s)", "(individual statement content)", [...new Set(statements.map(s=>s.channel||"email"))].join(","), statements.length, sent, failed, new Date().toISOString()]);
  return { ok:true, sent, failed };
}
/* Message your own team (staff Users), not members — by Email, SMS or
   WhatsApp, to everyone or specific people. Distinct from handleSendMessage_,
   which targets members. */
function handleSendTeamMessage_(body){
  const caller = verifySession_(body);
  if(!caller.isAdmin && !(canMessage_(caller))) throw new Error("You don't have permission to message your team.");
  const { subject, message, userIds, allUsers, channels } = body;
  if(!message) throw new Error("Message body is required.");
  if(!channels || !channels.length) throw new Error("Choose at least one channel (Email, SMS, or WhatsApp).");
  let targets = readRows_("Users").filter(u=> u.orgId === caller.orgId && u.status !== "Suspended");
  if(!allUsers){
    const idSet = {}; (userIds || []).forEach(id=> idSet[id] = true);
    targets = targets.filter(u=> idSet[u.userId]);
  }
  if(!targets.length) throw new Error("No matching team members to message.");
  let sent = 0, failed = 0;
  targets.forEach(u=>{
    const personalized = String(message).replace(/\{name\}/g, u.fullName || u.username || "there");
    if(channels.indexOf("email") > -1){
      if(u.email){ try{ MailApp.sendEmail(u.email, subject || "Message from your organization", personalized); sent++; }catch(e){ failed++; } }
      else failed++;
    }
    if(channels.indexOf("sms") > -1){
      if(u.phone && sendViaChannel_(u.phone, personalized, "sms")) sent++; else failed++;
    }
    if(channels.indexOf("whatsapp") > -1){
      if(u.phone && sendViaChannel_(u.phone, personalized, "whatsapp")) sent++; else failed++;
    }
  });
  sheet_("Messages").appendRow([Utilities.getUuid().slice(0,8).toUpperCase(), caller.orgId, caller.userId,
    "[Team] " + (subject || ""), message, channels.join(","), targets.length, sent, failed, new Date().toISOString()]);
  return { ok:true, recipients: targets.length, sent, failed };
}
function handleListMessages_(body){
  const caller = verifySession_(body);
  if(!canMessage_(caller)) throw new Error("You don't have permission to view messages.");
  const rows = readRows_("Messages").filter(m=> m.orgId === caller.orgId)
    .sort((a,b)=> String(b.sentAt).localeCompare(String(a.sentAt))).slice(0, 50);
  return { ok:true, messages: rows.map(m=> ({ subject:m.subject, channels:m.channels,
    recipientCount:m.recipientCount, sent:m.sent, failed:m.failed, sentAt:m.sentAt })) };
}

/* ================= trial/subscription email reminders =================
   Set up a daily time-driven trigger for this function (Apps Script editor
   > Triggers > Add Trigger > checkSubscriptionReminders > Time-driven >
   Day timer) so Admins get an email automatically as their trial or paid
   subscription nears/reaches its due date — the dashboard bell/banner
   already covers the in-app side of this in real time. */
/* Set up ONE daily time-driven trigger for this function (Apps Script editor
   > Triggers > Add Trigger > checkSubscriptionReminders > Time-driven > Day
   timer). It only actually sends anything on the 1st and 15th of each month
   (REMINDER_SEND_DAYS_OF_MONTH), so orgs get at most two reminder emails a
   month — never a daily nag — and only while their trial/subscription is
   genuinely within REMINDER_THRESHOLD_DAYS of ending. Each email also
   suggests their next plan/tier so it doubles as a light upsell nudge. */
function checkSubscriptionReminders(){
  const dayOfMonth = new Date().getDate();
  if(REMINDER_SEND_DAYS_OF_MONTH.indexOf(dayOfMonth) === -1) return;

  const rows = readRows_("Users").filter(u=> isTrue_(u.isAdmin) && !isTrue_(u.isSuperAdmin) && u.subscriptionExpiry);
  rows.forEach(u=>{
    const days = Math.ceil((new Date(u.subscriptionExpiry) - new Date()) / (1000*60*60*24));
    if(days > REMINDER_THRESHOLD_DAYS) return;

    // Reset the per-cycle counter whenever the expiry date itself has changed (renewed/upgraded).
    if(String(u.reminderCycleExpiry) !== String(u.subscriptionExpiry)){
      updateUserRow_(u.__row, { reminderCycleExpiry: u.subscriptionExpiry, remindersSentThisCycle: 0 });
      u.remindersSentThisCycle = 0;
    }
    const sentThisCycle = Number(u.remindersSentThisCycle) || 0;
    if(sentThisCycle >= 2) return; // already sent both reminders for this billing cycle

    const label = isTrue_(u.onTrial) ? "free trial" : "subscription";
    const subject = days <= 0
      ? `Your OKV SACCO ${label} has ended`
      : `Your OKV SACCO ${label} ${days === 1 ? "expires tomorrow" : `expires in ${days} days`}`;
    const nextTierNote = nextTierSuggestion_(u.plan);
    const bodyIntro = days <= 0
      ? `Hello ${u.fullName || u.username},\n\nYour ${label} for ${u.orgName || u.orgId} has ended. ` +
        `Sign in and open the Subscription tab to renew by bank transfer — you'll be confirmed and back up ` +
        `within the hour.`
      : `Hello ${u.fullName || u.username},\n\nYour ${label} for ${u.orgName || u.orgId} ` +
        `${days === 1 ? "expires tomorrow" : `expires in ${days} days`}. Renew any time from your dashboard's ` +
        `Subscription tab to avoid any interruption.`;
    const body = bodyIntro + (nextTierNote ? `\n\n${nextTierNote}` : "") +
      `\n\nReach us any time if you need help.\n\n— OKV Technology Consults`;
    try{ sendOkvEmail_(u.email || u.username, subject, body); }catch(e){ /* best-effort */ }
    updateUserRow_(u.__row, { lastReminderSent: new Date().toISOString(), remindersSentThisCycle: sentThisCycle + 1 });
  });
}
/* Suggests the next logical plan/tier for an email nudge — Basic → matching
   Professional tier, or if already Professional, the next-longer billing
   cycle for a lower effective monthly rate. */
function nextTierSuggestion_(currentPlanId){
  const plans = readRows_("Plans").filter(p=> isTrue_(p.active));
  if(!plans.length || !currentPlanId) return "";
  const current = plans.find(p=> p.planId === currentPlanId);
  if(!current) return "";
  let next = null;
  if(String(current.planId).indexOf("basic-") === 0){
    next = plans.find(p=> p.planId === String(current.planId).replace("basic-", "professional-"));
  } else {
    const order = ["monthly", "biannual", "yearly"];
    const idx = order.indexOf(current.cycleKey);
    if(idx > -1 && idx < order.length - 1){
      next = plans.find(p=> p.name === current.name && p.cycleKey === order[idx + 1]);
    }
  }
  if(!next) return "";
  return `Since you're on ${current.name} (${current.cycle}), you might also like ${next.name} (${next.cycle}) ` +
    `at ₦${Number(next.price).toLocaleString()} — see the Subscription tab for the full comparison.`;
}

/* ================= Super Admin: organizations, payments, announcements ================= */
function requireSuperAdmin_(caller){
  if(!isTrue_(caller.isSuperAdmin)) throw new Error("Only the Super Admin can do this.");
}
function handleListOrganizations_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const admins = readRows_("Users").filter(u=> isTrue_(u.isAdmin) && !isTrue_(u.isSuperAdmin));
  return { ok:true, organizations: admins.map(u=> ({
    orgId:u.orgId, orgName:u.orgName||"", adminUserId:u.userId, adminUsername:u.username,
    adminFullName:u.fullName||"", adminEmail:u.email||"", adminPhone:u.phone||"",
    plan:u.plan||"", billingCycle:u.billingCycle||"", subscriptionStatus:u.subscriptionStatus||"",
    subscriptionExpiry:u.subscriptionExpiry||"", onTrial:isTrue_(u.onTrial), status:u.status||"",
    createdAt:u.createdAt||""
  })) };
}
function handleUpdateOrganization_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const target = findUserRow_(u=> u.orgId === body.orgId && isTrue_(u.isAdmin));
  if(!target) throw new Error("Organization not found.");
  const updates = {};
  ["orgName","fullName","email","phone","plan","billingCycle","subscriptionStatus","subscriptionExpiry","status"]
    .forEach(k=>{ if(body[k] !== undefined) updates[k === "orgName" ? "orgName" : k] = body[k]; });
  if(body.onTrial !== undefined) updates.onTrial = !!body.onTrial;
  updateUserRow_(target.__row, updates);
  return { ok:true };
}
function handleListAllPayments_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const rows = readRows_("Payments").sort((a,b)=>{
    if(a.status === "Pending" && b.status !== "Pending") return -1;
    if(b.status === "Pending" && a.status !== "Pending") return 1;
    return String(b.submittedAt).localeCompare(String(a.submittedAt));
  });
  return { ok:true, payments: rows.map(p=> ({ paymentId:p.paymentId, orgId:p.orgId, orgName:p.orgName,
    fullName:p.fullName, email:p.email, phone:p.phone, planId:p.planId, planName:p.planName, amount:p.amount,
    proofUrl:p.proofUrl, status:p.status, submittedAt:p.submittedAt, reviewedAt:p.reviewedAt })) };
}
function handleReviewPayment_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const rows = readRows_("Payments");
  const payment = rows.find(p=> p.paymentId === body.paymentId);
  if(!payment) throw new Error("Payment not found.");
  const status = body.status; // "Confirmed" | "Rejected"
  updateSheetCell_("Payments", payment.__row, "status", status);
  updateSheetCell_("Payments", payment.__row, "reviewedAt", new Date().toISOString());
  const user = findUserRow_(u=> u.userId === payment.userId && u.orgId === payment.orgId);
  if(user){
    if(status === "Confirmed"){
      const days = planDurationDays_(payment.planId);
      const expiry = new Date(Date.now() + days*24*60*60*1000);
      updateUserRow_(user.__row, { subscriptionStatus:"active", subscriptionExpiry: expiry.toISOString(),
        plan: payment.planId, onTrial:false });
    }
    try{
      sendOkvEmail_(payment.email || user.email || user.username,
        "Your OKV SACCO subscription payment — " + status,
        status === "Confirmed"
          ? "Good news! Your payment has been confirmed and your subscription is now active.\nSign in to your dashboard to see your new plan and renewal date.\n\n— OKV Technology Consults"
          : "We could not confirm your payment. Please reach out to us on WhatsApp or email so we can help resolve this.\n\n— OKV Technology Consults");
    }catch(e){ /* best-effort */ }
  }
  updateSheetCell_("Payments", payment.__row, "notified", true);
  return { ok:true };
}
function handleSuperSendMessage_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const { subject, message, allOrgs, orgIds, channels } = body;
  if(!message) throw new Error("Message body is required.");
  if(!channels || !channels.length) throw new Error("Choose at least one channel (Email, SMS, or WhatsApp).");
  let admins = readRows_("Users").filter(u=> isTrue_(u.isAdmin) && !isTrue_(u.isSuperAdmin));
  if(!allOrgs){
    const idSet = {}; (orgIds || []).forEach(id=> idSet[id] = true);
    admins = admins.filter(u=> idSet[u.orgId]);
  }
  if(!admins.length) throw new Error("No matching organizations to message.");
  let sent = 0, failed = 0;
  admins.forEach(u=>{
    const personalized = String(message)
      .replace(/\{orgName\}/g, u.orgName || u.orgId)
      .replace(/\{adminName\}/g, u.fullName || u.username);
    if(channels.indexOf("email") > -1){
      if(u.email){ try{ sendOkvEmail_(u.email, subject || "Message from OKV Technology Consults", personalized); sent++; }catch(e){ failed++; } }
      else failed++;
    }
    if(channels.indexOf("sms") > -1){
      if(u.phone && sendViaChannel_(u.phone, personalized, "sms")) sent++; else failed++;
    }
    if(channels.indexOf("whatsapp") > -1){
      if(u.phone && sendViaChannel_(u.phone, personalized, "whatsapp")) sent++; else failed++;
    }
  });
  sheet_("AdminMessages").appendRow([Utilities.getUuid().slice(0,8).toUpperCase(), caller.username,
    subject || "", message, channels.join(","), admins.length, sent, failed, new Date().toISOString()]);
  return { ok:true, recipients: admins.length, sent, failed };
}
function handleListAdminMessages_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const rows = readRows_("AdminMessages").sort((a,b)=> String(b.sentAt).localeCompare(String(a.sentAt))).slice(0, 50);
  return { ok:true, messages: rows.map(m=> ({ subject:m.subject, channels:m.channels,
    recipientCount:m.recipientCount, sent:m.sent, failed:m.failed, sentAt:m.sentAt })) };
}

/* ================= platform settings (contact & bank details) =================
   Public read (marketing pages and the payment form need these before/without
   a session) — write is Super Admin only, editable from the Settings tab. */
function handleGetSettings_(body){
  const rows = readRows_("Settings");
  const map = {}; rows.forEach(r=> map[r.key] = r.value);
  Object.keys(DEFAULT_SETTINGS).forEach(k=>{ if(map[k] === undefined) map[k] = DEFAULT_SETTINGS[k]; });
  return {
    ok:true,
    contact: { whatsapp: map.contactWhatsapp, phone: map.contactPhone, email: map.contactEmail, website: map.contactWebsite },
    bank: { accountNumber: map.bankAccountNumber, bankName: map.bankName, accountName: map.bankAccountName },
    notifyEmail: map.notifyEmail,
    manualPaymentEnabled: String(map.manualPaymentEnabled) === "true",
    branding: { primaryColor: map.brandPrimaryColor, secondaryColor: map.brandSecondaryColor,
      logoUrl: map.brandLogoUrl, systemName: map.systemName, tagline: map.tagline },
    messaging: { smsProvider: map.smsProvider, smsSenderId: map.smsSenderId, smsConfigured: !!map.smsApiKey,
      whatsappProvider: map.whatsappProvider, whatsappSenderId: map.whatsappSenderId, whatsappConfigured: !!map.whatsappApiKey },
    gateways: {
      paystack: { enabled: String(map.paystackEnabled)==="true", publicKey: map.paystackPublicKey, configured: !!map.paystackSecretKey },
      remita: { enabled: String(map.remitaEnabled)==="true", publicKey: map.remitaPublicKey, configured: !!map.remitaSecretKey },
      flutterwave: { enabled: String(map.flutterwaveEnabled)==="true", publicKey: map.flutterwavePublicKey, configured: !!map.flutterwaveSecretKey }
    },
    trialDays: Number(map.trialDays) || 7
  };
}
/* Super Admin only — returns secret keys too, for populating the edit form.
   (handleGetSettings_ above deliberately omits secrets from the public response.) */
function handleGetSettingsFull_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const rows = readRows_("Settings");
  const map = {}; rows.forEach(r=> map[r.key] = r.value);
  Object.keys(DEFAULT_SETTINGS).forEach(k=>{ if(map[k] === undefined) map[k] = DEFAULT_SETTINGS[k]; });
  return Object.assign({ ok:true }, map);
}
function handleUpdateSettings_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const sh = sheet_("Settings");
  const rows = readRows_("Settings");
  const fieldMap = {
    whatsapp:"contactWhatsapp", phone:"contactPhone", email:"contactEmail", website:"contactWebsite",
    bankAccountNumber:"bankAccountNumber", bankName:"bankName", bankAccountName:"bankAccountName",
    notifyEmail:"notifyEmail", manualPaymentEnabled:"manualPaymentEnabled",
    brandPrimaryColor:"brandPrimaryColor", brandSecondaryColor:"brandSecondaryColor", brandLogoUrl:"brandLogoUrl",
    systemName:"systemName", tagline:"tagline",
    smsProvider:"smsProvider", smsApiKey:"smsApiKey", smsSenderId:"smsSenderId",
    whatsappProvider:"whatsappProvider", whatsappApiKey:"whatsappApiKey", whatsappSenderId:"whatsappSenderId",
    paystackEnabled:"paystackEnabled", paystackPublicKey:"paystackPublicKey", paystackSecretKey:"paystackSecretKey",
    remitaEnabled:"remitaEnabled", remitaPublicKey:"remitaPublicKey", remitaSecretKey:"remitaSecretKey",
    flutterwaveEnabled:"flutterwaveEnabled", flutterwavePublicKey:"flutterwavePublicKey", flutterwaveSecretKey:"flutterwaveSecretKey",
    trialDays:"trialDays"
  };
  Object.keys(fieldMap).forEach(inputKey=>{
    if(body[inputKey] === undefined) return;
    const settingKey = fieldMap[inputKey];
    const existing = rows.find(r=> r.key === settingKey);
    if(existing) updateSheetCell_("Settings", existing.__row, "value", body[inputKey]);
    else sh.appendRow([settingKey, body[inputKey]]);
  });
  return { ok:true };
}
function handleUploadLogo_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  if(!body.logoBase64) throw new Error("No logo file provided.");
  const folder = getOrCreateFolder_("OKV SACCO Branding");
  const commaIdx = body.logoBase64.indexOf(",");
  const raw = commaIdx > -1 ? body.logoBase64.slice(commaIdx+1) : body.logoBase64;
  const blob = Utilities.newBlob(Utilities.base64Decode(raw), "image/png", body.logoFileName || "logo.png");
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = "https://drive.google.com/uc?export=view&id=" + file.getId();
  const sh = sheet_("Settings");
  const existing = readRows_("Settings").find(r=> r.key === "brandLogoUrl");
  if(existing) updateSheetCell_("Settings", existing.__row, "value", url);
  else sh.appendRow(["brandLogoUrl", url]);
  return { ok:true, url };
}

/* ================= Plans & Pricing (Super Admin controlled) ================= */
function planPublicShape_(p){
  return { id:p.planId, name:p.name, cycle:p.cycle, cycleKey:p.cycleKey, price:Number(p.price)||0,
    audience:p.audience, features: JSON.parse(p.features || "[]") };
}
function handleListPlans_(body){
  const rows = readRows_("Plans").filter(p=> isTrue_(p.active))
    .sort((a,b)=> (Number(a.sortOrder)||0) - (Number(b.sortOrder)||0));
  const byCycle = { monthly:[], biannual:[], yearly:[] };
  rows.forEach(p=>{ if(byCycle[p.cycleKey]) byCycle[p.cycleKey].push(planPublicShape_(p)); });
  return { ok:true, plans: byCycle };
}
function handleListAllPlans_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const rows = readRows_("Plans").sort((a,b)=> (Number(a.sortOrder)||0) - (Number(b.sortOrder)||0));
  return { ok:true, plans: rows.map(p=> Object.assign(planPublicShape_(p), { active: isTrue_(p.active) })) };
}
function handleSavePlan_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const { id, name, cycle, cycleKey, price, audience, features, active, sortOrder } = body;
  if(!id || !name || !cycleKey) throw new Error("Plan ID, name and billing cycle are required.");
  const sh = sheet_("Plans");
  const rows = readRows_("Plans");
  const existing = rows.find(p=> p.planId === id);
  // Preserve the existing sortOrder unless the caller explicitly sent a new one —
  // the Super Admin edit form doesn't have a sortOrder field, so it must never
  // be silently reset to 0 just because a plan was edited.
  const resolvedSortOrder = sortOrder !== undefined ? Number(sortOrder) || 0
    : (existing ? Number(existing.sortOrder) || 0 : 0);
  const rowValues = [id, name, cycle || cycleKey, cycleKey, Number(price)||0, audience || "",
    JSON.stringify(features || []), active !== false, resolvedSortOrder];
  if(existing) sh.getRange(existing.__row, 1, 1, PLANS_HEADERS.length).setValues([rowValues]);
  else sh.appendRow(rowValues);
  return { ok:true };
}
function handleDeletePlan_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const rows = readRows_("Plans");
  const existing = rows.find(p=> p.planId === body.id);
  if(!existing) throw new Error("Plan not found.");
  sheet_("Plans").deleteRow(existing.__row);
  return { ok:true };
}

/* ================= tenant spreadsheet cell-capacity monitoring =================
   Google Sheets caps every spreadsheet FILE at 10,000,000 cells total,
   summed across all its tabs. Since each tenant's growing business records
   now live in their own spreadsheet (see getOrgSpreadsheet_ above), this
   tracks each one's usage against that cap so nobody hits it by surprise.

   Set up ONE time-driven trigger for checkAllTenantCellCapacity() —
   Apps Script editor > Triggers > Add Trigger > checkAllTenantCellCapacity
   > Time-driven > Day timer (or Week timer). It must NOT run on every page
   load: it writes its results into the master CellCapacity sheet, and the
   Super Admin dashboard's Capacity tab only ever reads that stored snapshot
   (via getCapacityReport), never recomputes live. */
function checkAllTenantCellCapacity(){
  const tenants = readRows_("Tenants");
  if(!tenants.length) return;
  const existingCap = readRows_("CellCapacity");
  const capByOrg = {}; existingCap.forEach(c=> capByOrg[c.orgId] = c);
  const capSheet = sheet_("CellCapacity");
  const now = new Date();
  const alerts = [];

  tenants.forEach(t=>{
    if(!t.spreadsheetId || t.status === "Archived") return;
    let ss;
    try{ ss = SpreadsheetApp.openById(t.spreadsheetId); }
    catch(err){ return; } // spreadsheet missing/inaccessible — skip rather than fail the whole run

    const totalCells = computeSpreadsheetCellUsage_(ss);
    const percentUsed = totalCells / CELL_CAPACITY_LIMIT;
    const prev = capByOrg[t.orgId];

    let history = [];
    if(prev && prev.historySnapshotJson){
      try{ history = JSON.parse(prev.historySnapshotJson); }catch(err){ history = []; }
    }
    history.push({ date: now.toISOString(), cells: totalCells });
    if(history.length > 100) history = history.slice(history.length - 100); // ~a season of daily snapshots is plenty

    const growth30 = estimateCellGrowth_(history, 30);
    const growth90 = estimateCellGrowth_(history, 90);
    const cols = DATA_HEADERS.length;
    const rowsAdded30d = Math.round(growth30.deltaCells / cols);
    const rowsAdded90d = Math.round(growth90.deltaCells / cols);
    // prefer the tighter 30-day window once there's enough history in it; fall back to 90-day
    const growthCellsPerDay = growth30.days >= 3 ? (growth30.deltaCells / growth30.days)
      : (growth90.days > 0 ? (growth90.deltaCells / growth90.days) : 0);

    let estDaysTo90Percent = "";
    const targetCells = CELL_CAPACITY_LIMIT * CELL_CAPACITY_PROJECTION_TARGET;
    if(totalCells >= targetCells) estDaysTo90Percent = 0;
    else if(growthCellsPerDay > 0) estDaysTo90Percent = Math.ceil((targetCells - totalCells) / growthCellsPerDay);
    // else: no measurable growth yet — leave blank rather than claim a false "never"

    let statusFlag = "Healthy";
    let recommendedAction = "No action needed — usage is well within limits.";
    if(percentUsed >= CELL_CAPACITY_ACTION_THRESHOLD){
      statusFlag = "Action Needed";
      recommendedAction = "Archive older records into a separate archive spreadsheet, or split this " +
        "tenant's data (e.g. by year) into an additional linked spreadsheet.";
    } else if(percentUsed >= CELL_CAPACITY_MONITOR_THRESHOLD){
      statusFlag = "Monitor";
      recommendedAction = "Usage is climbing — plan an archive or data split before this reaches 85%.";
    }

    const rowValues = [t.orgId, t.orgName || "", t.spreadsheetId, totalCells,
      Math.round(percentUsed * 10000) / 100, rowsAdded30d, rowsAdded90d,
      Math.round(growthCellsPerDay * 100) / 100, estDaysTo90Percent, statusFlag,
      recommendedAction, now.toISOString(), JSON.stringify(history)];

    if(prev) capSheet.getRange(prev.__row, 1, 1, CELL_CAPACITY_HEADERS.length).setValues([rowValues]);
    else capSheet.appendRow(rowValues);

    if(statusFlag === "Action Needed" && (!prev || prev.statusFlag !== "Action Needed")){
      alerts.push({ orgName: t.orgName || t.orgId, percentUsed: rowValues[4], estDaysTo90Percent });
    }
  });

  if(alerts.length) notifySuperAdminCapacityAlerts_(alerts);
}
/* Sums getLastRow() x getLastColumn() across every tab actually in use —
   the real used range, not each tab's full default 1000x26 grid. */
function computeSpreadsheetCellUsage_(ss){
  let total = 0;
  ss.getSheets().forEach(sh=>{
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if(lastRow > 0 && lastCol > 0) total += lastRow * lastCol;
  });
  return total;
}
/* Looks at the stored daily snapshots within the last `windowDays` and
   returns the cell growth across that span. */
function estimateCellGrowth_(history, windowDays){
  const cutoff = Date.now() - windowDays*24*60*60*1000;
  const inWindow = history.filter(h=> new Date(h.date).getTime() >= cutoff);
  if(inWindow.length < 2) return { deltaCells: 0, days: 0 };
  const first = inWindow[0], last = inWindow[inWindow.length - 1];
  const days = (new Date(last.date) - new Date(first.date)) / (24*60*60*1000);
  return { deltaCells: Math.max(0, last.cells - first.cells), days: days || 0 };
}
function notifySuperAdminCapacityAlerts_(alerts){
  const to = getSetting_("notifyEmail");
  if(!to) return;
  const lines = alerts.map(a=> "- " + a.orgName + ": " + a.percentUsed + "% of the cell limit used" +
    (a.estDaysTo90Percent !== "" ? (", est. " + a.estDaysTo90Percent + " day(s) to 90%") : "")).join("\n");
  try{
    sendOkvEmail_(to, "OKV SACCO — tenant(s) crossed the Action Needed capacity threshold",
      "The following organizations have crossed 85% of their spreadsheet's 10,000,000-cell limit:\n\n" +
      lines + "\n\nReview them in the Super Admin dashboard's Capacity tab.\n\n— OKV SACCO System");
  }catch(err){ /* best-effort */ }
}
/* Super Admin only — reads the stored snapshot from CellCapacity, never
   recomputes live (that's checkAllTenantCellCapacity()'s job, on a trigger). */
function handleGetCapacityReport_(body){
  const caller = verifySession_(body); requireSuperAdmin_(caller);
  const rows = readRows_("CellCapacity").sort((a,b)=> (Number(b.percentUsed)||0) - (Number(a.percentUsed)||0));
  return { ok:true, capacity: rows.map(r=> ({
    orgId: r.orgId, orgName: r.orgName || r.orgId, percentUsed: Number(r.percentUsed) || 0,
    totalCellsUsed: Number(r.totalCellsUsed) || 0, rowsAdded30d: Number(r.rowsAdded30d) || 0,
    rowsAdded90d: Number(r.rowsAdded90d) || 0, estDaysTo90Percent: r.estDaysTo90Percent,
    statusFlag: r.statusFlag || "Healthy", recommendedAction: r.recommendedAction || "",
    lastChecked: r.lastChecked || "" })) };
}

/* ================= ONE-TIME MIGRATION: single-sheet -> per-tenant =================
   Run this ONCE, manually, from the Apps Script editor (select
   migrateExistingDataToPerTenant in the function dropdown, click Run) if
   this spreadsheet already has business records sitting in a single shared
   "Data" tab from before the per-tenant architecture existed. It:
     1) Makes sure Tenants/CellCapacity exist (calls setupSheets()).
     2) For every distinct orgId found in Users that doesn't already have a
        Tenants entry, creates a fresh per-tenant spreadsheet, copies that
        org's rows out of the old shared "Data" tab into it, and registers
        the new spreadsheet ID in the Tenants sheet.
     3) Renames (never deletes) the old shared "Data" tab so nothing is
        lost if anything looks off afterwards.
   Safe to re-run: any orgId that already has a Tenants row is skipped, so
   a second run only picks up orgs that signed up since the last run. */
function migrateExistingDataToPerTenant(){
  setupSheets();
  const oldDataSheet = SS.getSheetByName("Data");
  if(!oldDataSheet){ Logger.log("No legacy shared 'Data' sheet found on this spreadsheet — nothing to migrate."); return; }

  const alreadyMigrated = {}; readRows_("Tenants").forEach(t=> alreadyMigrated[t.orgId] = true);
  const admins = readRows_("Users").filter(u=> isTrue_(u.isAdmin) && !isTrue_(u.isSuperAdmin));
  const orgsById = {}; admins.forEach(u=>{ if(!orgsById[u.orgId]) orgsById[u.orgId] = u; });
  const oldRows = readRows_("Data");

  let migratedOrgCount = 0, migratedRowCount = 0;
  Object.keys(orgsById).forEach(orgId=>{
    if(alreadyMigrated[orgId]){ Logger.log("Skipping " + orgId + " — already has a tenant spreadsheet."); return; }
    const admin = orgsById[orgId];
    const tenantSS = createTenantSpreadsheet_(orgId, admin.orgName || orgId);
    const rows = oldRows.filter(r=> r.orgId === orgId);
    if(rows.length){
      const values = rows.map(r=> DATA_HEADERS.map(h=> r[h]));
      sheet_("Data", tenantSS).getRange(2, 1, values.length, DATA_HEADERS.length).setValues(values);
    }
    migratedOrgCount++; migratedRowCount += rows.length;
    Logger.log("Migrated " + rows.length + " record(s) for " + orgId + " -> " + tenantSS.getUrl());
  });

  if(migratedOrgCount > 0){
    oldDataSheet.setName("Data_ARCHIVED_PRE_MIGRATION_" + Date.now());
    Logger.log("Migration complete: " + migratedOrgCount + " organization(s), " + migratedRowCount +
      " record(s) total. The old shared Data tab was renamed (not deleted) for safety.");
  } else {
    Logger.log("Nothing new to migrate — every organization already has its own tenant spreadsheet.");
  }
}
