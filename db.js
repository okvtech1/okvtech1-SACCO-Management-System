/* OKV SACCO — Subscribe (Online) — local storage layer.
   Everything is mirrored into IndexedDB so the app keeps working offline;
   syncPush/syncPull (in dashboard.html) reconcile this with the Google
   Sheet backend whenever a connection is available. */

const LDB_NAME = "OKVSubscribeDB";
const LDB_VERSION = 1;
const RECORD_STORES = ["members","contributions","loans","repayments","dividends"];
let _ldb = null;

function openLocalDB(){
  return new Promise((resolve, reject)=>{
    const req = indexedDB.open(LDB_NAME, LDB_VERSION);
    req.onupgradeneeded = (e)=>{
      const db = e.target.result;
      RECORD_STORES.forEach(name=>{
        if(!db.objectStoreNames.contains(name)) db.createObjectStore(name, {keyPath:"recordId"});
      });
      if(!db.objectStoreNames.contains("session")) db.createObjectStore("session", {keyPath:"key"});
      if(!db.objectStoreNames.contains("meta")) db.createObjectStore("meta", {keyPath:"key"});
    };
    req.onsuccess = (e)=>{ _ldb = e.target.result; resolve(_ldb); };
    req.onerror = (e)=> reject(e.target.error);
  });
}
function ldbGetAll(store){
  return new Promise((resolve, reject)=>{
    const tx = _ldb.transaction(store, "readonly");
    const req = tx.objectStore(store).getAll();
    req.onsuccess = ()=> resolve(req.result || []);
    req.onerror = (e)=> reject(e.target.error);
  });
}
function ldbPut(store, item){
  return new Promise((resolve, reject)=>{
    const tx = _ldb.transaction(store, "readwrite");
    tx.objectStore(store).put(item);
    tx.oncomplete = ()=> resolve();
    tx.onerror = (e)=> reject(e.target.error);
  });
}
function ldbGet(store, key){
  return new Promise((resolve, reject)=>{
    const tx = _ldb.transaction(store, "readonly");
    const req = tx.objectStore(store).get(key);
    req.onsuccess = ()=> resolve(req.result || null);
    req.onerror = (e)=> reject(e.target.error);
  });
}

/* ---- session (who is logged in on this device) ---- */
async function saveSession(session){
  await openLocalDB();
  await ldbPut("session", Object.assign({key:"current"}, session, {savedAt: new Date().toISOString()}));
}
async function loadSession(){
  await openLocalDB();
  return await ldbGet("session", "current");
}
async function clearSession(){
  await openLocalDB();
  await ldbPut("session", {key:"current", cleared:true});
}

/* ---- last successful sync timestamp ---- */
async function getLastSync(){
  await openLocalDB();
  const row = await ldbGet("meta", "sync");
  return row ? row.since : null;
}
async function setLastSync(iso){
  await ldbPut("meta", {key:"sync", since: iso});
}
