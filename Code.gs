// ============================================================
//  ระบบติดตามยาสมุนไพร สสจ.ตาก — v2
//  3 Roles | Multi-Year | Dynamic Products
// ============================================================

const SPREADSHEET_ID = '1RD7ceCMEHsFWkmIePyvBrA_gRJdCSphY8L4Jq3NawmU';

const SH = {
  USERS:'users', DISTRICTS:'districts', PRODUCTS:'products',
  FY:'fiscal_years', BUDGETS:'district_budgets',
  DEMAND:'demand_plans', PRODUCTION:'production_records',
  DISTRIBUTION:'distribution_records',
  AUDIT:'audit_log', NOTIF:'notifications',
};

const ROLE = { ADMIN:'admin', DISTRICT:'district', PHHO:'phho' };

// ---- ENTRY POINTS ------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบยาสมุนไพร สสจ.ตาก')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleRequest(raw) {
  try {
    const { action, token, data } = JSON.parse(raw);
    if (action === 'login') return out(login(data));
    const sess = verifySession(token);
    if (!sess.ok) return out({ ok:false, error:'กรุณาเข้าสู่ระบบใหม่' });
    const u = sess.user;
    switch (action) {
      case 'logout':              return out(logout(token));
      case 'changePassword':      return out(changePassword(data, u));
      // Lookups
      case 'getProducts':         return out(getProducts(data));
      case 'getDistricts':        return out(getDistricts(u));
      case 'getFiscalYears':      return out(getFiscalYears());
      case 'getCurrentFY':        return out(getCurrentFY());
      // Products CRUD (admin)
      case 'saveProduct':         return out(adminOnly(u, () => saveProduct(data, u)));
      case 'toggleProduct':       return out(adminOnly(u, () => toggleProduct(data, u)));
      // Districts CRUD (admin)
      case 'saveDistrict':        return out(adminOnly(u, () => saveDistrict(data, u)));
      case 'toggleDistrict':      return out(adminOnly(u, () => toggleDistrict(data, u)));
      // Fiscal Years (admin)
      case 'saveFiscalYear':      return out(adminOnly(u, () => saveFiscalYear(data, u)));
      case 'setCurrentFY':        return out(adminOnly(u, () => setCurrentFY(data, u)));
      // Budgets
      case 'getBudgets':          return out(getBudgets(data, u));
      case 'saveBudget':          return out(adminOnly(u, () => saveBudget(data, u)));
      // Demand
      case 'getDemandPlans':      return out(getDemandPlans(data, u));
      case 'saveDemandPlan':      return out(distOrAdmin(u, () => saveDemandPlan(data, u)));
      case 'submitDemandPlan':    return out(distOrAdmin(u, () => submitDemandPlan(data, u)));
      case 'reviewDemandPlan':    return out(adminOnly(u, () => reviewDemandPlan(data, u)));
      // Production (admin)
      case 'getProduction':       return out(getProduction(data));
      case 'saveProduction':      return out(adminOnly(u, () => saveProduction(data, u)));
      // Distribution
      case 'getDistribution':     return out(getDistribution(data, u));
      case 'createDistribution':  return out(adminOnly(u, () => createDistribution(data, u)));
      case 'approveDistribution': return out(adminOnly(u, () => approveDistribution(data, u)));
      // Reports
      case 'getDashboard':        return out(getDashboard(data, u));
      case 'getBudgetTracker':    return out(getBudgetTracker(data, u));
      case 'getMonthlyReport':    return out(getMonthlyReport(data, u));
      // Notifications
      case 'getNotifications':    return out(getNotifications(u));
      case 'markRead':            return out(markRead(data, u));
      // Users (admin)
      case 'getUsers':            return out(adminOnly(u, () => getUsers()));
      case 'saveUser':            return out(adminOnly(u, () => saveUser(data, u)));
      case 'toggleUser':          return out(adminOnly(u, () => toggleUser(data, u)));
      case 'resetPassword':       return out(adminOnly(u, () => resetPassword(data, u)));
      default: return out({ ok:false, error:'Unknown action: ' + action });
    }
  } catch(e) {
    Logger.log('handleRequest error: ' + e.message);
    return JSON.stringify({ ok:false, error:e.message });
  }
}

function out(o)               { return JSON.stringify(o); }
function adminOnly(u, fn)     { return u.role===ROLE.ADMIN ? fn() : noAccess(); }
function distOrAdmin(u, fn)   { return [ROLE.ADMIN,ROLE.DISTRICT].includes(u.role) ? fn() : noAccess(); }
function noAccess()           { return { ok:false, error:'ไม่มีสิทธิ์ดำเนินการนี้' }; }

// ---- SHEET HELPERS -----------------------------------------
function ss()        { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sh(name)    {
  const s = ss().getSheetByName(name);
  if (!s) throw new Error('Sheet not found: ' + name);
  return s;
}
function rows(name) {
  const s = sh(name), d = s.getDataRange().getValues();
  if (d.length < 2) return [];
  const h = d[0];
  return d.slice(1).map(r => { const o={}; h.forEach((k,i)=>o[k]=r[i]); return o; });
}
function append(name, obj) {
  const s = sh(name), h = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  s.appendRow(h.map(k => obj[k]!==undefined ? obj[k] : ''));
}
function updateById(name, idField, idVal, upd) {
  const s = sh(name), d = s.getDataRange().getValues(), h = d[0];
  const ic = h.indexOf(idField);
  for (let i=1;i<d.length;i++) {
    if (String(d[i][ic])===String(idVal)) {
      Object.keys(upd).forEach(k=>{const c=h.indexOf(k);if(c>=0)s.getRange(i+1,c+1).setValue(upd[k]);});
      return true;
    }
  }
  return false;
}
function nextId(name, idField) {
  const s = sh(name); if(s.getLastRow()<2) return 1;
  const ic = s.getRange(1,1,1,s.getLastColumn()).getValues()[0].indexOf(idField);
  const ids = s.getRange(2,ic+1,s.getLastRow()-1,1).getValues().flat().map(v=>parseInt(v)||0);
  return Math.max(0,...ids)+1;
}
const now = () => new Date().toISOString();

// ---- SESSION -----------------------------------------------
function login({ username, password }) {
  if (!username||!password) return { ok:false, error:'กรุณากรอกข้อมูลให้ครบ' };
  const user = rows(SH.USERS).find(u =>
    u.username===username && u.password_hash===hashPw(password) && u.is_active==true
  );
  if (!user) return { ok:false, error:'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  const token = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('S_'+token, JSON.stringify({
    user_id:user.user_id, username:user.username, full_name:user.full_name,
    role:user.role, district_id:user.district_id,
    exp: Date.now()+(8*3600*1000),
  }));
  updateById(SH.USERS,'user_id',user.user_id,{last_login:now()});
  audit(user.user_id,'users','LOGIN',user.user_id,'','login');
  return {
    ok:true, token,
    user:{ user_id:user.user_id, full_name:user.full_name, role:user.role, district_id:user.district_id },
  };
}

function logout(token) {
  PropertiesService.getScriptProperties().deleteProperty('S_'+token);
  return { ok:true };
}

function verifySession(token) {
  if (!token) return { ok:false };
  const raw = PropertiesService.getScriptProperties().getProperty('S_'+token);
  if (!raw) return { ok:false };
  const s = JSON.parse(raw);
  if (Date.now()>s.exp) { PropertiesService.getScriptProperties().deleteProperty('S_'+token); return { ok:false }; }
  return { ok:true, user:s };
}

function hashPw(pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8)
    .map(b=>('0'+(b&0xff).toString(16)).slice(-2)).join('');
}

function changePassword({ old_password, new_password }, u) {
  if (!old_password||!new_password) return { ok:false, error:'กรุณากรอกข้อมูลให้ครบ' };
  if (new_password.length<8) return { ok:false, error:'รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร' };
  const user = rows(SH.USERS).find(x=>x.user_id==u.user_id);
  if (!user||user.password_hash!==hashPw(old_password)) return { ok:false, error:'รหัสผ่านเดิมไม่ถูกต้อง' };
  updateById(SH.USERS,'user_id',u.user_id,{password_hash:hashPw(new_password)});
  audit(u.user_id,'users','CHANGE_PASSWORD',u.user_id,'','');
  return { ok:true };
}

// ---- FISCAL YEARS ------------------------------------------
function getFiscalYears() {
  return { ok:true, data:rows(SH.FY).filter(y=>y.is_active!=false) };
}
function getCurrentFY() {
  const all = rows(SH.FY).filter(y=>y.is_active!=false);
  return { ok:true, data: all.find(y=>y.is_current==true) || all[all.length-1] || null };
}
function saveFiscalYear({ fy_id, fiscal_year, description }, u) {
  if (!fiscal_year) return { ok:false, error:'กรุณาระบุปีงบประมาณ' };
  if (fy_id) {
    updateById(SH.FY,'fy_id',fy_id,{fiscal_year,description:description||''});
    audit(u.user_id,SH.FY,'UPDATE',fy_id,'',String(fiscal_year));
    return { ok:true, fy_id };
  }
  const id = nextId(SH.FY,'fy_id');
  append(SH.FY,{fy_id:id,fiscal_year,description:description||'',is_active:true,is_current:false,created_at:now()});
  audit(u.user_id,SH.FY,'CREATE',id,'',String(fiscal_year));
  return { ok:true, fy_id:id };
}
function setCurrentFY({ fy_id }, u) {
  rows(SH.FY).forEach(y=>updateById(SH.FY,'fy_id',y.fy_id,{is_current:y.fy_id==fy_id}));
  audit(u.user_id,SH.FY,'SET_CURRENT',fy_id,'','');
  return { ok:true };
}

// ---- PRODUCTS ----------------------------------------------
function getProducts({ include_inactive }={}) {
  let p = rows(SH.PRODUCTS);
  if (!include_inactive) p = p.filter(x=>x.is_active!=false);
  return { ok:true, data:p };
}
function saveProduct({ product_id, product_name, product_code, unit, price_per_unit, category, description }, u) {
  if (!product_name||!unit||!price_per_unit) return { ok:false, error:'กรุณากรอกข้อมูลสำคัญให้ครบ' };
  const ts=now(), price=parseFloat(price_per_unit);
  if (product_id) {
    const old = rows(SH.PRODUCTS).find(p=>p.product_id==product_id);
    updateById(SH.PRODUCTS,'product_id',product_id,{
      product_name,product_code:product_code||'',unit,price_per_unit:price,
      category:category||'',description:description||'',updated_at:ts,updated_by:u.user_id
    });
    audit(u.user_id,SH.PRODUCTS,'UPDATE',product_id,JSON.stringify(old),`price=${price}`);
    return { ok:true, product_id };
  }
  const id = nextId(SH.PRODUCTS,'product_id');
  append(SH.PRODUCTS,{product_id:id,product_name,product_code:product_code||'',unit,price_per_unit:price,
    category:category||'',description:description||'',is_active:true,
    created_at:ts,created_by:u.user_id,updated_at:ts,updated_by:u.user_id});
  audit(u.user_id,SH.PRODUCTS,'CREATE',id,'',product_name);
  return { ok:true, product_id:id };
}
function toggleProduct({ product_id, is_active }, u) {
  const old = rows(SH.PRODUCTS).find(p=>p.product_id==product_id);
  updateById(SH.PRODUCTS,'product_id',product_id,{is_active,updated_at:now(),updated_by:u.user_id});
  audit(u.user_id,SH.PRODUCTS,is_active?'ACTIVATE':'DEACTIVATE',product_id,JSON.stringify(old),'');
  return { ok:true };
}

// ---- DISTRICTS ---------------------------------------------
function getDistricts(u) {
  let d = rows(SH.DISTRICTS).filter(x=>x.is_active!=false);
  if (u.role===ROLE.DISTRICT) d = d.filter(x=>x.district_id==u.district_id);
  return { ok:true, data:d };
}
function saveDistrict({ district_id, district_name, district_code }, u) {
  if (!district_name) return { ok:false, error:'กรุณาระบุชื่ออำเภอ' };
  if (district_id) {
    updateById(SH.DISTRICTS,'district_id',district_id,{district_name,district_code:district_code||''});
    audit(u.user_id,SH.DISTRICTS,'UPDATE',district_id,'',district_name);
    return { ok:true, district_id };
  }
  const id = nextId(SH.DISTRICTS,'district_id');
  append(SH.DISTRICTS,{district_id:id,district_name,district_code:district_code||'',is_active:true});
  audit(u.user_id,SH.DISTRICTS,'CREATE',id,'',district_name);
  return { ok:true, district_id:id };
}
function toggleDistrict({ district_id, is_active }, u) {
  updateById(SH.DISTRICTS,'district_id',district_id,{is_active});
  audit(u.user_id,SH.DISTRICTS,is_active?'ACTIVATE':'DEACTIVATE',district_id,'','');
  return { ok:true };
}

// ---- BUDGETS -----------------------------------------------
function getBudgets({ fy_id }, u) {
  let b = rows(SH.BUDGETS);
  if (fy_id) b = b.filter(x=>x.fy_id==fy_id);
  if (u.role===ROLE.DISTRICT) b = b.filter(x=>x.district_id==u.district_id);
  return { ok:true, data:b };
}
function saveBudget({ budget_id, district_id, fy_id, annual_quota_thb, remarks }, u) {
  if (!district_id||!fy_id||!annual_quota_thb) return { ok:false, error:'ข้อมูลไม่ครบ' };
  const quota = parseFloat(String(annual_quota_thb).replace(/,/g,''));
  const upd   = {annual_quota_thb:quota,remarks:remarks||'',approved_by:u.full_name,approved_at:now()};
  if (budget_id) {
    updateById(SH.BUDGETS,'budget_id',budget_id,upd);
    audit(u.user_id,SH.BUDGETS,'UPDATE',budget_id,'',`quota=${quota}`);
    return { ok:true, budget_id };
  }
  const dup = rows(SH.BUDGETS).find(b=>b.district_id==district_id&&b.fy_id==fy_id);
  if (dup) {
    updateById(SH.BUDGETS,'budget_id',dup.budget_id,upd);
    audit(u.user_id,SH.BUDGETS,'UPDATE',dup.budget_id,'',`quota=${quota}`);
    return { ok:true, budget_id:dup.budget_id };
  }
  const id = nextId(SH.BUDGETS,'budget_id');
  append(SH.BUDGETS,{budget_id:id,district_id,fy_id,annual_quota_thb:quota,
    approved_by:u.full_name,approved_at:now(),remarks:remarks||''});
  audit(u.user_id,SH.BUDGETS,'CREATE',id,'',`d=${district_id} fy=${fy_id} q=${quota}`);
  return { ok:true, budget_id:id };
}

// ---- DEMAND PLANNING ---------------------------------------
function getDemandPlans({ fy_id, quarter, district_id:did }, u) {
  let p = rows(SH.DEMAND);
  if (fy_id)   p = p.filter(x=>x.fy_id==fy_id);
  if (quarter) p = p.filter(x=>x.quarter==quarter);
  if (u.role===ROLE.DISTRICT) p = p.filter(x=>x.district_id==u.district_id);
  else if (did) p = p.filter(x=>x.district_id==did);
  return { ok:true, data:p };
}
function saveDemandPlan({ plan_id, district_id, product_id, fy_id, quarter, quantity_requested, remarks }, u) {
  if (u.role===ROLE.DISTRICT&&u.district_id!=district_id) return noAccess();
  if (!district_id||!product_id||!fy_id||!quarter||!quantity_requested) return { ok:false, error:'ข้อมูลไม่ครบ' };
  const prod = rows(SH.PRODUCTS).find(p=>p.product_id==product_id);
  if (!prod) return { ok:false, error:'ไม่พบผลิตภัณฑ์' };
  const price = parseFloat(prod.price_per_unit);
  const total = quantity_requested * price;
  if (plan_id) {
    const ex = rows(SH.DEMAND).find(p=>p.plan_id==plan_id);
    if (ex&&ex.status==='approved') return { ok:false, error:'ไม่สามารถแก้ไขแผนที่อนุมัติแล้ว' };
    updateById(SH.DEMAND,'plan_id',plan_id,{quantity_requested,unit_price_snapshot:price,total_value:total,remarks:remarks||'',status:'draft'});
    audit(u.user_id,SH.DEMAND,'UPDATE',plan_id,'',`qty=${quantity_requested}`);
    return { ok:true, plan_id };
  }
  const id = nextId(SH.DEMAND,'plan_id');
  append(SH.DEMAND,{plan_id:id,district_id,product_id,fy_id,quarter,
    quantity_requested,unit_price_snapshot:price,total_value:total,
    status:'draft',remarks:remarks||'',
    submitted_at:'',submitted_by:'',reviewed_at:'',reviewed_by:'',review_note:''});
  audit(u.user_id,SH.DEMAND,'CREATE',id,'',`d=${district_id} p=${product_id} qty=${quantity_requested}`);
  return { ok:true, plan_id:id };
}
function submitDemandPlan({ plan_ids }, u) {
  if (!plan_ids||!plan_ids.length) return { ok:false, error:'ไม่มีรายการ' };
  const ts = now();
  plan_ids.forEach(pid=>{
    const p = rows(SH.DEMAND).find(x=>x.plan_id==pid);
    if (!p||p.status==='approved') return;
    if (u.role===ROLE.DISTRICT&&p.district_id!=u.district_id) return;
    updateById(SH.DEMAND,'plan_id',pid,{status:'submitted',submitted_at:ts,submitted_by:u.user_id});
  });
  notifyAdmins('demand_submitted','แผนความต้องการใหม่รออนุมัติ',
    `${u.full_name} ส่งแผน ${plan_ids.length} รายการ`,SH.DEMAND,plan_ids[0]);
  audit(u.user_id,SH.DEMAND,'SUBMIT','','',plan_ids.join(','));
  return { ok:true };
}
function reviewDemandPlan({ plan_id, approved, review_note }, u) {
  const status = approved?'approved':'rejected';
  const old = rows(SH.DEMAND).find(p=>p.plan_id==plan_id);
  updateById(SH.DEMAND,'plan_id',plan_id,{status,reviewed_at:now(),reviewed_by:u.user_id,review_note:review_note||''});
  if (old&&old.submitted_by)
    notify(old.submitted_by,'demand_'+status,
      `แผนของคุณ${approved?'ได้รับการอนุมัติ':'ถูกปฏิเสธ'}`,review_note||'',SH.DEMAND,plan_id);
  audit(u.user_id,SH.DEMAND,status.toUpperCase(),plan_id,'',review_note||'');
  return { ok:true };
}

// ---- PRODUCTION --------------------------------------------
function getProduction({ fy_id, month, product_id }) {
  let r = rows(SH.PRODUCTION);
  if (fy_id)      r = r.filter(x=>x.fy_id==fy_id);
  if (month)      r = r.filter(x=>x.month==month);
  if (product_id) r = r.filter(x=>x.product_id==product_id);
  return { ok:true, data:r };
}
function saveProduction({ record_id, product_id, fy_id, month, plan_quantity, actual_quantity, notes }, u) {
  if (!product_id||!fy_id||!month) return { ok:false, error:'ข้อมูลไม่ครบ' };
  const prod = rows(SH.PRODUCTS).find(p=>p.product_id==product_id);
  if (!prod) return { ok:false, error:'ไม่พบผลิตภัณฑ์' };
  const price = parseFloat(prod.price_per_unit);
  const pv=plan_quantity*price, av=actual_quantity*price, ts=now();
  const upd={plan_quantity,actual_quantity,plan_value:pv,actual_value:av,notes:notes||'',recorded_at:ts,recorded_by:u.user_id};
  if (record_id) {
    updateById(SH.PRODUCTION,'record_id',record_id,upd);
  } else {
    const ex = rows(SH.PRODUCTION).find(r=>r.product_id==product_id&&r.fy_id==fy_id&&r.month==month);
    if (ex) updateById(SH.PRODUCTION,'record_id',ex.record_id,upd);
    else { const id=nextId(SH.PRODUCTION,'record_id'); append(SH.PRODUCTION,{record_id:id,product_id,fy_id,month,...upd}); }
  }
  if (plan_quantity>0&&actual_quantity/plan_quantity<0.8)
    notifyAdmins('production_low','ผลผลิตต่ำกว่าแผน',
      `${prod.product_name} เดือน ${month}: ${actual_quantity}/${plan_quantity} หน่วย`,SH.PRODUCTION,null);
  audit(u.user_id,SH.PRODUCTION,'SAVE','','',`p=${product_id} fy=${fy_id} m=${month} plan=${plan_quantity} actual=${actual_quantity}`);
  return { ok:true };
}

// ---- DISTRIBUTION ------------------------------------------
function getDistribution({ fy_id, district_id:did, status:st }, u) {
  let r = rows(SH.DISTRIBUTION);
  if (fy_id) r = r.filter(x=>x.fy_id==fy_id);
  if (st)    r = r.filter(x=>x.status===st);
  if (u.role===ROLE.DISTRICT) r = r.filter(x=>x.district_id==u.district_id);
  else if (did) r = r.filter(x=>x.district_id==did);
  return { ok:true, data:r };
}
function createDistribution({ district_id, product_id, fy_id, quantity_issued, plan_id, issued_ref }, u) {
  if (!district_id||!product_id||!fy_id||!quantity_issued) return { ok:false, error:'ข้อมูลไม่ครบ' };
  const qc = checkQuota(district_id,fy_id,quantity_issued,product_id);
  if (!qc.ok) return qc;
  const prod  = rows(SH.PRODUCTS).find(p=>p.product_id==product_id);
  const price = prod ? parseFloat(prod.price_per_unit) : 0;
  const id    = nextId(SH.DISTRIBUTION,'dist_id');
  append(SH.DISTRIBUTION,{dist_id:id,district_id,product_id,plan_id:plan_id||'',fy_id,
    quantity_issued,unit_price_snapshot:price,total_value:quantity_issued*price,
    status:'pending',requested_at:now(),requested_by:u.user_id,
    approved_at:'',approved_by:'',issued_at:'',issued_ref:issued_ref||''});
  audit(u.user_id,SH.DISTRIBUTION,'CREATE',id,'',`d=${district_id} p=${product_id} qty=${quantity_issued}`);
  return { ok:true, dist_id:id };
}
function approveDistribution({ dist_id, approved, issued_ref }, u) {
  const status = approved?'issued':'rejected';
  updateById(SH.DISTRIBUTION,'dist_id',dist_id,{
    status,approved_at:now(),approved_by:u.user_id,
    issued_at:approved?now():'',issued_ref:issued_ref||''});
  audit(u.user_id,SH.DISTRIBUTION,status.toUpperCase(),dist_id,'','');
  return { ok:true };
}
function checkQuota(district_id, fy_id, qty_new, product_id) {
  const budget = rows(SH.BUDGETS).find(b=>b.district_id==district_id&&b.fy_id==fy_id);
  if (!budget) return { ok:false, error:'ยังไม่มีการตั้งงบสำหรับอำเภอและปีงบนี้' };
  const used = rows(SH.DISTRIBUTION)
    .filter(r=>r.district_id==district_id&&r.fy_id==fy_id&&r.status==='issued')
    .reduce((s,r)=>s+parseFloat(r.total_value||0),0);
  const prod  = rows(SH.PRODUCTS).find(p=>p.product_id==product_id);
  const price = prod ? parseFloat(prod.price_per_unit) : 0;
  const cost  = qty_new * price;
  if (used+cost > parseFloat(budget.annual_quota_thb))
    return { ok:false, error:`โควต้าไม่เพียงพอ (ใช้ไป ${fmtN(used)} / ${fmtN(budget.annual_quota_thb)} บาท)` };
  return { ok:true, remaining:budget.annual_quota_thb-used-cost };
}

// ---- DASHBOARD & REPORTS -----------------------------------
function getDashboard({ fy_id }, u) {
  const fyr = fy_id || (getCurrentFY().data||{}).fy_id;
  const budgets  = rows(SH.BUDGETS).filter(b=>b.fy_id==fyr);
  const districts= rows(SH.DISTRICTS).filter(d=>d.is_active!=false);
  const distrib  = rows(SH.DISTRIBUTION).filter(r=>r.fy_id==fyr&&r.status==='issued');
  const demands  = rows(SH.DEMAND).filter(p=>p.fy_id==fyr);
  const prods    = rows(SH.PRODUCTS).filter(p=>p.is_active!=false);
  const scope    = u.role===ROLE.DISTRICT ? districts.filter(d=>d.district_id==u.district_id) : districts;
  const totalQuota = budgets.filter(b=>scope.some(d=>d.district_id==b.district_id)).reduce((s,b)=>s+parseFloat(b.annual_quota_thb||0),0);
  const totalUsed  = distrib.filter(r=>scope.some(d=>d.district_id==r.district_id)).reduce((s,r)=>s+parseFloat(r.total_value||0),0);
  const districtSummary = scope.map(d=>{
    const budget = budgets.find(b=>b.district_id==d.district_id);
    const quota  = budget ? parseFloat(budget.annual_quota_thb) : 0;
    const used   = distrib.filter(r=>r.district_id==d.district_id).reduce((s,r)=>s+parseFloat(r.total_value||0),0);
    return { district_id:d.district_id, district_name:d.district_name, quota, used,
      remaining:quota-used, pct:quota?Math.round(used/quota*100):0,
      plan_submitted:demands.some(p=>p.district_id==d.district_id&&['submitted','approved'].includes(p.status)) };
  });
  return { ok:true, data:{
    fy_id:fyr, total_quota:totalQuota, total_used:totalUsed,
    remaining:totalQuota-totalUsed, pct:totalQuota?Math.round(totalUsed/totalQuota*100):0,
    district_summary:districtSummary,
    not_submitted:districtSummary.filter(d=>!d.plan_submitted).map(d=>d.district_name),
    pending_dist:rows(SH.DISTRIBUTION).filter(r=>r.fy_id==fyr&&r.status==='pending').length,
    product_count:prods.length,
  }};
}

function getBudgetTracker({ fy_id }, u) {
  const fyr = fy_id || (getCurrentFY().data||{}).fy_id;
  const budgets  = rows(SH.BUDGETS).filter(b=>b.fy_id==fyr);
  const districts= rows(SH.DISTRICTS).filter(d=>d.is_active!=false);
  const distrib  = rows(SH.DISTRIBUTION).filter(r=>r.fy_id==fyr&&r.status==='issued');
  const prods    = rows(SH.PRODUCTS).filter(p=>p.is_active!=false);
  const scope    = u.role===ROLE.DISTRICT ? districts.filter(d=>d.district_id==u.district_id) : districts;
  const data = scope.map(d=>{
    const budget = budgets.find(b=>b.district_id==d.district_id);
    const quota  = budget ? parseFloat(budget.annual_quota_thb) : 0;
    const byProd = prods.map(p=>{
      const qty = distrib.filter(r=>r.district_id==d.district_id&&r.product_id==p.product_id)
        .reduce((s,r)=>s+parseInt(r.quantity_issued||0),0);
      return { product_id:p.product_id, product_name:p.product_name, unit:p.unit,
        price:parseFloat(p.price_per_unit), qty, val:qty*parseFloat(p.price_per_unit) };
    });
    const used = byProd.reduce((s,b)=>s+b.val,0);
    return { district_id:d.district_id, district_name:d.district_name,
      quota, used, remaining:quota-used, pct:quota?Math.round(used/quota*100):0, by_product:byProd };
  });
  return { ok:true, data };
}

function getMonthlyReport({ fy_id, month }, u) {
  const fyr = fy_id || (getCurrentFY().data||{}).fy_id;
  const districts= rows(SH.DISTRICTS).filter(d=>d.is_active!=false);
  const prods    = rows(SH.PRODUCTS).filter(p=>p.is_active!=false);
  const distrib  = rows(SH.DISTRIBUTION).filter(r=>r.status==='issued'&&r.fy_id==fyr&&
    (!month||new Date(r.issued_at||r.requested_at).getMonth()+1==month));
  const allDistrib= rows(SH.DISTRIBUTION).filter(r=>r.status==='issued'&&r.fy_id==fyr);
  const production= rows(SH.PRODUCTION).filter(r=>r.fy_id==fyr&&(!month||r.month==month));
  const demands   = rows(SH.DEMAND).filter(r=>r.fy_id==fyr&&r.status!=='rejected');
  const scope     = u.role===ROLE.DISTRICT ? districts.filter(d=>d.district_id==u.district_id) : districts;

  const reportA = scope.map(d=>{
    const row={district_name:d.district_name,total:0,by_product:{}};
    prods.forEach(p=>{
      const qty=distrib.filter(r=>r.district_id==d.district_id&&r.product_id==p.product_id)
        .reduce((s,r)=>s+parseInt(r.quantity_issued||0),0);
      const val=qty*parseFloat(p.price_per_unit);
      row.by_product[p.product_id]={qty,val,unit:p.unit}; row.total+=val;
    });
    return row;
  }).filter(r=>r.total>0);

  const reportB = scope.map(d=>{
    const row={district_name:d.district_name,total_acc:0,by_product:{}};
    prods.forEach(p=>{
      const qty=allDistrib.filter(r=>r.district_id==d.district_id&&r.product_id==p.product_id)
        .reduce((s,r)=>s+parseInt(r.quantity_issued||0),0);
      const val=qty*parseFloat(p.price_per_unit);
      row.by_product[p.product_id]={qty,val,unit:p.unit}; row.total_acc+=val;
    });
    return row;
  }).filter(r=>r.total_acc>0);

  const reportC = prods.map(p=>{
    const prod=production.find(r=>r.product_id==p.product_id)||{};
    const totalReq=demands.filter(d=>d.product_id==p.product_id).reduce((s,d)=>s+parseInt(d.quantity_requested||0),0);
    const pln=parseInt(prod.plan_quantity||0), act=parseInt(prod.actual_quantity||0);
    return { product_id:p.product_id, product_name:p.product_name, unit:p.unit,
      plan:pln, actual:act, requested:totalReq, pct:pln?Math.round(act/pln*100):0, meets_demand:act>=totalReq };
  });

  return { ok:true, data:{reportA,reportB,reportC,fy_id:fyr,month} };
}

// ---- USERS -------------------------------------------------
function getUsers() {
  return { ok:true, data:rows(SH.USERS).map(u=>({...u,password_hash:'***'})) };
}
function saveUser({ user_id, username, password, role, district_id, full_name, email, phone }, u) {
  if (!username||!role) return { ok:false, error:'ข้อมูลไม่ครบ' };
  if (role===ROLE.DISTRICT&&!district_id) return { ok:false, error:'กรุณาระบุอำเภอสำหรับ District User' };
  if (user_id) {
    const upd={role,district_id:district_id||'',full_name:full_name||'',email:email||'',phone:phone||''};
    if (password) upd.password_hash = hashPw(password);
    updateById(SH.USERS,'user_id',user_id,upd);
    audit(u.user_id,SH.USERS,'UPDATE',user_id,'',`role=${role}`);
    return { ok:true, user_id };
  }
  if (!password) return { ok:false, error:'กรุณาตั้งรหัสผ่าน' };
  if (rows(SH.USERS).find(x=>x.username===username)) return { ok:false, error:'ชื่อผู้ใช้นี้มีอยู่แล้ว' };
  const id=nextId(SH.USERS,'user_id');
  append(SH.USERS,{user_id:id,username,password_hash:hashPw(password),
    role,district_id:district_id||'',full_name:full_name||'',email:email||'',phone:phone||'',
    is_active:true,last_login:'',created_at:now(),created_by:u.user_id});
  audit(u.user_id,SH.USERS,'CREATE',id,'',`${username} role=${role}`);
  return { ok:true, user_id:id };
}
function toggleUser({ user_id, is_active }, u) {
  if (user_id==u.user_id) return { ok:false, error:'ไม่สามารถระงับบัญชีตัวเองได้' };
  updateById(SH.USERS,'user_id',user_id,{is_active});
  audit(u.user_id,SH.USERS,is_active?'ACTIVATE':'DEACTIVATE',user_id,'','');
  return { ok:true };
}
function resetPassword({ user_id, new_password }, u) {
  if (!new_password||new_password.length<8) return { ok:false, error:'รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร' };
  updateById(SH.USERS,'user_id',user_id,{password_hash:hashPw(new_password)});
  audit(u.user_id,SH.USERS,'RESET_PASSWORD',user_id,'','');
  return { ok:true };
}

// ---- NOTIFICATIONS -----------------------------------------
function getNotifications(u) {
  return { ok:true, data:rows(SH.NOTIF)
    .filter(x=>x.user_id==u.user_id)
    .sort((a,b)=>new Date(b.created_at)-new Date(a.created_at)).slice(0,40) };
}
function markRead({ notif_id }, u) {
  updateById(SH.NOTIF,'notif_id',notif_id,{is_read:true});
  return { ok:true };
}
function notify(user_id, type, title, message, refTable, refId) {
  const id=nextId(SH.NOTIF,'notif_id');
  append(SH.NOTIF,{notif_id:id,user_id,type,title,message:message||'',is_read:false,
    created_at:now(),ref_table:refTable||'',ref_id:refId||''});
}
function notifyAdmins(type, title, message, refTable, refId) {
  rows(SH.USERS).filter(u=>u.role===ROLE.ADMIN&&u.is_active==true)
    .forEach(u=>notify(u.user_id,type,title,message,refTable,refId));
}

// ---- AUDIT -------------------------------------------------
function audit(user_id, table, action, record_id, old_val, new_val) {
  try {
    const id=nextId(SH.AUDIT,'log_id');
    append(SH.AUDIT,{log_id:id,user_id,table_name:table,action,record_id:record_id||'',
      old_value:old_val||'',new_value:new_val||'',logged_at:now()});
  } catch(e){ Logger.log('Audit fail: '+e.message); }
}
function fmtN(n){ return Math.round(parseFloat(n)||0).toLocaleString(); }

// ---- TIME TRIGGERS -----------------------------------------
function checkDeadlines() {
  const fy=getCurrentFY().data; if(!fy) return;
  const dl=new Date(new Date().getFullYear(),3,17);
  const days=Math.ceil((dl-new Date())/86400000);
  if(![1,3,7].includes(days)) return;
  const submitted=new Set(rows(SH.DEMAND)
    .filter(p=>p.fy_id==fy.fy_id&&['submitted','approved'].includes(p.status)).map(p=>p.district_id));
  rows(SH.DISTRICTS).filter(d=>d.is_active!=false&&!submitted.has(d.district_id)).forEach(d=>{
    rows(SH.USERS).filter(u=>u.district_id==d.district_id&&u.is_active==true).forEach(u=>{
      notify(u.user_id,'deadline_warning',`กำหนดส่งแผนอีก ${days} วัน`,
        `กรุณาส่งแผนความต้องการยาสมุนไพร ปีงบ ${fy.fiscal_year} ก่อนครบกำหนด`,SH.DEMAND,null);
    });
  });
}
function checkLowQuota() {
  const fy=getCurrentFY().data; if(!fy) return;
  rows(SH.BUDGETS).filter(b=>b.fy_id==fy.fy_id).forEach(b=>{
    const used=rows(SH.DISTRIBUTION).filter(r=>r.district_id==b.district_id&&r.fy_id==fy.fy_id&&r.status==='issued')
      .reduce((s,r)=>s+parseFloat(r.total_value||0),0);
    const quota=parseFloat(b.annual_quota_thb);
    if(quota>0&&used/quota>=0.8) {
      const d=rows(SH.DISTRICTS).find(x=>x.district_id==b.district_id);
      rows(SH.USERS).filter(u=>u.district_id==b.district_id&&u.is_active==true).forEach(u=>{
        notify(u.user_id,'quota_low','โควต้าคงเหลือน้อยกว่า 20%',
          `${d?d.district_name:''}: ใช้ไป ${fmtN(used)} / ${fmtN(quota)} บาท`,SH.BUDGETS,b.budget_id);
      });
    }
  });
}

// ---- SETUP -------------------------------------------------
const SCHEMAS = {
  users:                ['user_id','username','password_hash','role','district_id','full_name','email','phone','is_active','last_login','created_at','created_by'],
  districts:            ['district_id','district_name','district_code','is_active'],
  products:             ['product_id','product_name','product_code','unit','price_per_unit','category','description','is_active','created_at','created_by','updated_at','updated_by'],
  fiscal_years:         ['fy_id','fiscal_year','description','is_active','is_current','created_at'],
  district_budgets:     ['budget_id','district_id','fy_id','annual_quota_thb','approved_by','approved_at','remarks'],
  demand_plans:         ['plan_id','district_id','product_id','fy_id','quarter','quantity_requested','unit_price_snapshot','total_value','status','remarks','submitted_at','submitted_by','reviewed_at','reviewed_by','review_note'],
  production_records:   ['record_id','product_id','fy_id','month','plan_quantity','actual_quantity','plan_value','actual_value','notes','recorded_at','recorded_by'],
  distribution_records: ['dist_id','district_id','product_id','plan_id','fy_id','quantity_issued','unit_price_snapshot','total_value','status','requested_at','requested_by','approved_at','approved_by','issued_at','issued_ref'],
  audit_log:            ['log_id','user_id','table_name','action','record_id','old_value','new_value','logged_at'],
  notifications:        ['notif_id','user_id','type','title','message','is_read','created_at','ref_table','ref_id'],
};

function setupSpreadsheet() {
  const sp = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.entries(SCHEMAS).forEach(([name,headers])=>{
    let s=sp.getSheetByName(name);
    if(!s) s=sp.insertSheet(name);
    if(s.getLastRow()===0){
      s.appendRow(headers);
      s.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#E8E6F0');
    }
  });

  // Seed fiscal years
  const fyS=sp.getSheetByName('fiscal_years');
  if(fyS.getLastRow()<=1){
    fyS.appendRow([1,2568,'ปีงบประมาณ 2568',true,false,now()]);
    fyS.appendRow([2,2569,'ปีงบประมาณ 2569',true,true,now()]);
    fyS.appendRow([3,2570,'ปีงบประมาณ 2570',true,false,now()]);
  }

  // Seed districts
  const dS=sp.getSheetByName('districts');
  if(dS.getLastRow()<=1){
    [['เมืองตาก','D01'],['แม่สอด','D02'],['บ้านตาก','D03'],['สามเงา','D04'],
     ['แม่ระมาด','D05'],['ท่าสองยาง','D06'],['พบพระ','D07'],['อุ้มผาง','D08'],['วังเจ้า','D09']]
      .forEach((d,i)=>dS.appendRow([i+1,d[0],d[1],true]));
  }

  // Seed products
  const pS=sp.getSheetByName('products');
  if(pS.getLastRow()<=1){
    const ts=now();
    [[1,'ยาน้ำแก้อักเสบมะขามป้อม (120 ml.)','P001','ขวด',28,'ยาน้ำ',''],
     [2,'ยาธาตุอบเชย (120 ml.)','P002','ขวด',20,'ยาน้ำ',''],
     [3,'ขี้ผึ้งไพล 20 กรัม','P003','กระปุก',30,'ยาทาภายนอก','']]
      .forEach(p=>pS.appendRow([...p,true,ts,1,ts,1]));
  }

  // Seed budgets for FY2569 (fy_id=2)
  const bS=sp.getSheetByName('district_budgets');
  if(bS.getLastRow()<=1){
    const ts=now();
    [[1,370000],[2,400000],[3,150000],[4,100000],[5,200000],
     [6,300000],[7,240000],[8,120000],[9,120000]]
      .forEach((b,i)=>bS.appendRow([i+1,b[0],2,b[1],'System',ts,'งบประมาณปีงบ 2569']));
  }

  // Seed users
  const uS=sp.getSheetByName('users');
  if(uS.getLastRow()<=1){
    const ts=now();
    uS.appendRow([1,'admin',hashPw('Admin@2569'),ROLE.ADMIN,'','ผู้ดูแลระบบ รพ.บ้านตาก','admin@bantakhospital.go.th','',true,'',ts,1]);
    uS.appendRow([2,'phho',hashPw('Phho@2569'),ROLE.PHHO,'','กลุ่มงานแพทย์แผนไทยฯ สสจ.ตาก','phho@tak.moph.go.th','',true,'',ts,1]);
  }

  Logger.log('Setup สำเร็จ!\nadmin / Admin@2569\nphho  / Phho@2569');
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkDeadlines').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('checkLowQuota').timeBased().everyDays(1).atHour(9).create();
  Logger.log('Triggers ตั้งค่าสำเร็จ');
}
