/**
 * TAGADDOD - Google Apps Script backend
 * Bind this project to the destination Google Spreadsheet before deployment.
 */
const CONFIG = {
  SPREADSHEET_ID: '',
  PRIMARY_ADMIN_PHONE: '01276341407',
  DEFAULT_ALLOWANCE: 50,
  SESSION_DAYS: 30,
  RESULT_CACHE_SECONDS: 90,
  ACCOUNTS_SHEET: 'Accounts',
  TRIPS_SHEET: 'Trips',
  EXTRA_DAYS_SHEET: 'ExtraDays',
  SETTLEMENTS_SHEET: 'Settlements',
  SETTINGS_SHEET: 'Settings'
};

const ACCOUNT_HEADERS = ['id','name','phone','role','passwordSalt','passwordHash','avatar','createdAt'];
const TRIP_HEADERS = ['id','ownerPhone','repName','timestamp','date','place','client','allowance','createdAt','status','reviewedBy','reviewedAt'];
const EXTRA_HEADERS = ['id','ownerPhone','repName','timestamp','date','reason','proofUrl','createdAt','status','reviewedBy','reviewedAt'];
const SETTLEMENT_HEADERS = ['id','ownerPhone','repName','date','receivedToday','priorBalance','totalCash','expensesJson','spent','remaining','createdAt'];
const SETTING_HEADERS = ['key','value'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  let result;
  try {
    result = p.action === 'result' ? readCachedResult_(p.requestId) : {ok:false,error:'طلب غير معروف.'};
  } catch (err) {
    result = {ok:false,error:String(err && err.message ? err.message : err)};
  }
  return output_(result, p.callback);
}

function doPost(e) {
  let request = {};
  try {
    request = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    ensureSheets_();
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let result;
    try {
      const cached = request.requestId ? readCachedResult_(request.requestId) : {pending:true};
      result = cached && cached.pending !== true ? cached : handleAction_(request);
      if (request.requestId) writeCachedResult_(request.requestId,result);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    const result = {ok:false,error:String(err && err.message ? err.message : err)};
    if (request.requestId) writeCachedResult_(request.requestId, result);
  }
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

function handleAction_(request) {
  const action = String(request.action || '');
  const payload = request.payload || {};
  if (action === 'login') return login_(payload);

  const session = getSession_(request.sessionToken);
  if (!session) throw new Error('انتهت الجلسة. سجل الدخول مرة أخرى.');
  if (action === 'logout') {
    PropertiesService.getScriptProperties().deleteProperty(session.key);
    return {ok:true};
  }
  const actor = findAccountByPhone_(session.phone);
  if (!actor) throw new Error('الحساب غير موجود.');
  const admin = isAdmin_(actor);

  if (action === 'get_state') return {ok:true,data:getState_(session.phone,payload)};
  if (action === 'list_trips') {
    let trips = getTrips_().filter(t => admin || normalizePhone_(t.ownerPhone) === actor.phone);
    const from = String(payload.from || '');
    const to = String(payload.to || '');
    if (from) trips = trips.filter(t => t.date >= from);
    if (to) trips = trips.filter(t => t.date <= to);
    return {ok:true,trips};
  }
  if (action === 'list_extra_days') {
    let extra = getExtraDays_().filter(t => admin || normalizePhone_(t.ownerPhone) === actor.phone);
    const from = String(payload.from || '');
    const to = String(payload.to || '');
    if (from) extra = extra.filter(t => t.date >= from);
    if (to) extra = extra.filter(t => t.date <= to);
    return {ok:true,extraDays:extra};
  }
  if (action === 'review_extra_day') {
    requireAdmin_(admin);
    const status = String(payload.status || '');
    if (status !== 'approved' && status !== 'rejected') throw new Error('حالة المراجعة غير صحيحة.');
    reviewRecord_('extra',payload.id,status,actor);
    return {ok:true};
  }
  if (action === 'review_records') {
    requireAdmin_(admin);
    const type = String(payload.type || '');
    const status = String(payload.status || '');
    if (type !== 'extra') throw new Error('مراجعة الرحلات ملغاة؛ يمكن حذف الرحلة بدلًا من ذلك.');
    if (status !== 'approved' && status !== 'rejected') throw new Error('حالة المراجعة غير صحيحة.');
    const ids = Array.isArray(payload.ids) ? payload.ids : [];
    if (!ids.length || ids.length > 5000) throw new Error('حدد سجلات للمراجعة (بحد أقصى 5000 سجل في الطلب الواحد).');
    return {ok:true,...reviewPendingRecords_(type,ids,status,actor)};
  }

  switch (action) {
    case 'add_trip': {
      const owner = admin ? findAccountByPhone_(payload.repPhone) : actor;
      if (!owner) throw new Error('المندوب المحدد غير موجود.');
      const date = String(payload.date || '');
      const place = String(payload.place || '').trim();
      const client = String(payload.client || '').trim();
      const allowance = Number(getSetting_('defaultAllowance',CONFIG.DEFAULT_ALLOWANCE));
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !place || !client || !Number.isFinite(allowance) || allowance < 0) throw new Error('بيانات الرحلة غير مكتملة أو غير صحيحة.');
      const trips = getTrips_();
      const tripId = 'trip_' + safeId_(request.requestId || Utilities.getUuid());
      const existingTrip = trips.find(t => t.id === tripId);
      if (existingTrip) return {ok:true,tripId:existingTrip.id,duplicateRequest:true};
      if (trips.some(t => normalizePhone_(t.ownerPhone) === owner.phone && t.date === date)) throw new Error('تم تسجيل رحلة لهذا المندوب في هذا التاريخ بالفعل.');
      const trip = {id:tripId,ownerPhone:owner.phone,repName:owner.name,timestamp:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'),date,place,client,allowance,createdAt:new Date().toISOString(),status:'approved',reviewedBy:'',reviewedAt:''};
      getSheet_(CONFIG.TRIPS_SHEET,TRIP_HEADERS).appendRow(TRIP_HEADERS.map(k => trip[k] === undefined ? '' : trip[k]));
      return {ok:true,tripId:trip.id};
    }
    case 'delete_trip': {
      const sh = getSheet_(CONFIG.TRIPS_SHEET,TRIP_HEADERS);
      const rows = sh.getDataRange().getValues();
      const index = rows.findIndex((row,i) => i > 0 && String(row[0]) === String(payload.id));
      if (index < 1) throw new Error('الرحلة غير موجودة.');
      if (!admin && normalizePhone_(rows[index][1]) !== actor.phone) throw new Error('لا يمكنك حذف رحلة مندوب آخر.');
      sh.deleteRow(index + 1);
      return {ok:true};
    }
    case 'add_settlement': {
      const owner = admin ? findAccountByPhone_(payload.repPhone) : actor;
      if (!owner) throw new Error('المندوب المحدد غير موجود.');
      const date = String(payload.date || '');
      const receivedToday = Number(payload.receivedToday);
      const priorBalance = Number(payload.priorBalance);
      const rawExpenses = Array.isArray(payload.expenses) ? payload.expenses : [];
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('تاريخ التسوية غير صحيح.');
      if (!Number.isFinite(receivedToday) || receivedToday < 0 || !Number.isFinite(priorBalance) || priorBalance < 0) throw new Error('راجع المبلغ المستلم والرصيد السابق.');
      if (rawExpenses.length > 100) throw new Error('لا يمكن إضافة أكثر من 100 بند في التسوية.');
      const expenses = rawExpenses.map(item => {
        const name = String(item && item.name || '').trim();
        const amount = Number(item && item.amount);
        const gratuity = Number(item && item.gratuity || 0);
        if (!name || name.length > 150 || !Number.isFinite(amount) || amount < 0 || !Number.isFinite(gratuity) || gratuity < 0) throw new Error('راجع أسماء الجهات والمبالغ والإكراميات في بنود المصروفات.');
        return {name,amount,gratuity};
      });
      const settlementId = 'settlement_' + safeId_(request.requestId || Utilities.getUuid());
      const existing = getSettlements_().find(item => item.id === settlementId);
      if (existing) return {ok:true,settlementId:existing.id,duplicateRequest:true};
      const roundMoney = value => Math.round((value + Number.EPSILON) * 100) / 100;
      const totalCash = roundMoney(receivedToday + priorBalance);
      const spent = roundMoney(expenses.reduce((sum,item) => sum + item.amount + item.gratuity,0));
      const remaining = roundMoney(totalCash - spent);
      const settlement = {
        id:settlementId,
        ownerPhone:owner.phone,
        repName:owner.name,
        date,
        receivedToday:roundMoney(receivedToday),
        priorBalance:roundMoney(priorBalance),
        totalCash,
        expensesJson:JSON.stringify(expenses),
        spent,
        remaining,
        createdAt:new Date().toISOString()
      };
      getSheet_(CONFIG.SETTLEMENTS_SHEET,SETTLEMENT_HEADERS).appendRow(SETTLEMENT_HEADERS.map(key => settlement[key] === undefined ? '' : settlement[key]));
      pruneSettlements_();
      return {ok:true,settlementId:settlement.id};
    }
    case 'add_extra_day': {
      const owner = admin ? findAccountByPhone_(payload.repPhone) : actor;
      if (!owner) throw new Error('المندوب المحدد غير موجود.');
      const date = String(payload.date || '');
      const reason = String(payload.reason || '').trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !reason) throw new Error('بيانات اليوم الإضافي غير مكتملة.');
      const extraId = 'extra_' + safeId_(request.requestId || Utilities.getUuid());
      const existingExtra = getExtraDays_().find(item => item.id === extraId);
      if (existingExtra) return {ok:true,extraId:existingExtra.id,duplicateRequest:true};
      const proofUrl = saveProofImage_(payload.proofData,payload.proofName,payload.proofType);
      const extra = {id:extraId,ownerPhone:owner.phone,repName:owner.name,timestamp:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'),date,reason,proofUrl,createdAt:new Date().toISOString(),status:'pending',reviewedBy:'',reviewedAt:''};
      getSheet_(CONFIG.EXTRA_DAYS_SHEET,EXTRA_HEADERS).appendRow(EXTRA_HEADERS.map(k => extra[k] === undefined ? '' : extra[k]));
      return {ok:true,extraId:extra.id};
    }
    case 'delete_extra_day': {
      const sh = getSheet_(CONFIG.EXTRA_DAYS_SHEET,EXTRA_HEADERS);
      const rows = sh.getDataRange().getValues();
      const index = rows.findIndex((row,i) => i > 0 && String(row[0]) === String(payload.id));
      if (index < 1) throw new Error('السجل غير موجود.');
      if (!admin && normalizePhone_(rows[index][1]) !== actor.phone) throw new Error('لا يمكنك حذف سجل مندوب آخر.');
      sh.deleteRow(index + 1);
      return {ok:true};
    }
    case 'add_account': {
      requireAdmin_(admin);
      const name = String(payload.name || '').trim();
      const phone = normalizePhone_(payload.phone);
      const role = payload.role === 'admin' ? 'admin' : 'rep';
      if (!name || phone.length < 10) throw new Error('الاسم أو رقم الهاتف غير صحيح.');
      if (findAccountByPhone_(phone)) throw new Error('رقم الهاتف مسجل بالفعل.');
      const account = {id:(role==='admin'?'admin_':'rep_') + Utilities.getUuid(),name,phone,role,passwordSalt:'',passwordHash:'',avatar:'',createdAt:new Date().toISOString()};
      getSheet_(CONFIG.ACCOUNTS_SHEET,ACCOUNT_HEADERS).appendRow(ACCOUNT_HEADERS.map(k => account[k] || ''));
      return {ok:true};
    }
    case 'delete_account': {
      requireAdmin_(admin);
      const phone = normalizePhone_(payload.phone);
      const target = findAccountByPhone_(phone);
      if (!target || phone === actor.phone) throw new Error('لا يمكن حذف حسابك.');
      if (phone === CONFIG.PRIMARY_ADMIN_PHONE) throw new Error('لا يمكن حذف حساب المشرف علي رجب.');
      if (isAdmin_(target) && !isPrimaryAdmin_(actor)) throw new Error('المشرفون الآخرون لا يمكنهم حذف حساب مشرف.');
      const sh = getSheet_(CONFIG.ACCOUNTS_SHEET,ACCOUNT_HEADERS);
      const rows = sh.getDataRange().getValues();
      const index = rows.findIndex((row,i) => i > 0 && normalizePhone_(row[2]) === phone);
      if (index > 0) sh.deleteRow(index + 1);
      return {ok:true};
    }
    case 'reset_password': {
      requireAdmin_(admin);
      const phone = normalizePhone_(payload.phone);
      const target = findAccountByPhone_(phone);
      if (!target) throw new Error('الحساب غير موجود.');
      if (phone === CONFIG.PRIMARY_ADMIN_PHONE && !isPrimaryAdmin_(actor)) throw new Error('لا يمكنك إعادة ضبط كلمة سر المشرف علي رجب.');
      const sh = getSheet_(CONFIG.ACCOUNTS_SHEET,ACCOUNT_HEADERS);
      sh.getRange(target._row,5,1,2).setValues([['','']]);
      return {ok:true};
    }
    case 'set_allowance': {
      requireAdmin_(admin);
      const allowance = Number(payload.allowance);
      if (!Number.isFinite(allowance) || allowance < 0) throw new Error('قيمة البدل غير صحيحة.');
      setSetting_('defaultAllowance',allowance);
      return {ok:true};
    }
    case 'update_profile': {
      const sh = getSheet_(CONFIG.ACCOUNTS_SHEET,ACCOUNT_HEADERS);
      const account = findAccountByPhone_(actor.phone);
      const newPhone = normalizePhone_(payload.phone);
      const newPassword = String(payload.newPassword || '');
      const avatar = String(payload.avatar || '');
      if (newPhone && newPhone !== actor.phone) {
        if (newPhone.length < 10) throw new Error('رقم الهاتف الجديد غير صحيح.');
        if (findAccountByPhone_(newPhone)) throw new Error('رقم الهاتف الجديد مسجل بالفعل.');
      }
      let salt = account.passwordSalt;
      let hash = account.passwordHash;
      if (newPassword) {
        if (!verifyPassword_(String(payload.currentPassword || ''),salt,hash)) throw new Error('كلمة السر الحالية غير صحيحة.');
        if (newPassword.length < 4) throw new Error('كلمة السر الجديدة يجب أن تكون 4 خانات على الأقل.');
        salt = Utilities.getUuid() + Utilities.getUuid();
        hash = hashPassword_(newPassword,salt);
      }
      if (avatar && avatar.length > 45000) throw new Error('الصورة كبيرة جدًا. اختر صورة أصغر.');
      const updatedPhone = newPhone || actor.phone;
      sh.getRange(account._row,3).setValue(updatedPhone);
      sh.getRange(account._row,5,1,2).setValues([[salt,hash]]);
      if (avatar) sh.getRange(account._row,7).setValue(avatar);
      if (updatedPhone !== actor.phone) updateSessionPhone_(request.sessionToken,updatedPhone);
      return {ok:true};
    }
    default:
      throw new Error('عملية غير مدعومة.');
  }
}

function login_(payload) {
  const phone = normalizePhone_(payload.phone);
  const account = findAccountByPhone_(phone);
  if (!account) throw new Error('هذا الرقم غير مسجل. تواصل مع المشرف لإضافتك.');
  let salt = account.passwordSalt;
  let hash = account.passwordHash;
  if (!hash) {
    const newPassword = String(payload.newPassword || '');
    if (!newPassword) {
      if (!String(payload.password || '')) throw new Error('اكتب أي رقم سري للمتابعة.');
      return {ok:false,code:'FIRST_LOGIN',firstLogin:true,error:'هذا أول دخول للحساب. أنشئ رقمًا سريًا جديدًا.'};
    }
    if (!/^[0-9]{4,}$/.test(newPassword)) throw new Error('الرقم السري يجب أن يتكون من 4 أرقام على الأقل.');
    salt = Utilities.getUuid() + Utilities.getUuid();
    hash = hashPassword_(newPassword,salt);
    getSheet_(CONFIG.ACCOUNTS_SHEET,ACCOUNT_HEADERS).getRange(account._row,5,1,2).setValues([[salt,hash]]);
  } else if (!verifyPassword_(String(payload.password || ''),salt,hash)) {
    throw new Error('كلمة السر غير صحيحة.');
  }
  const token = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  PropertiesService.getScriptProperties().setProperty('SESSION_' + token,JSON.stringify({phone:account.phone,expiresAt:Date.now()+CONFIG.SESSION_DAYS*86400000}));
  return {ok:true,sessionToken:token};
}

function getState_(phone,payload) {
  const user = findAccountByPhone_(phone);
  if (!user) throw new Error('الحساب غير موجود.');
  const admin = isAdmin_(user);
  const page = Math.max(1,Number(payload.tripPage)||1);
  const extraPage = Math.max(1,Number(payload.extraPage)||1);
  const pageSize = 10;
  const allTrips = getTrips_();
  let trips = allTrips.filter(t => admin || normalizePhone_(t.ownerPhone) === user.phone);
  trips.sort((a,b) => String(b.date||'').localeCompare(String(a.date||'')) || String(b.timestamp||'').localeCompare(String(a.timestamp||'')));
  let extra = getExtraDays_().filter(t => admin || normalizePhone_(t.ownerPhone) === user.phone);
  extra.sort((a,b) => String(b.date||'').localeCompare(String(a.date||'')) || String(b.timestamp||'').localeCompare(String(a.timestamp||'')));
  const pendingExtraDays = extra.filter(t => t.status === 'pending');
  const accounts = admin ? getAccounts_() : [user];
  const summaryMonth = /^\d{4}-\d{2}$/.test(String(payload.summaryMonth || ''))
    ? String(payload.summaryMonth)
    : Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM');
  const monthTrips = trips.filter(trip => String(trip.date || '').slice(0,7) === summaryMonth);
  const monthlySummary = admin ? accounts
    .filter(account => account.role === 'rep' || monthTrips.some(trip => normalizePhone_(trip.ownerPhone) === account.phone))
    .map(account => {
      const ownTrips = monthTrips.filter(trip => normalizePhone_(trip.ownerPhone) === account.phone);
      const dates = {};
      ownTrips.forEach(trip => { dates[trip.date] = true; });
      return {
        repPhone:account.phone,
        repName:account.name,
        days:Object.keys(dates).length,
        allowance:ownTrips.reduce((sum,trip) => sum + Number(trip.allowance || 0),0)
      };
    }) : [];
  pruneSettlements_();
  const settlements = getSettlements_()
    .filter(item => admin || normalizePhone_(item.ownerPhone) === user.phone)
    .sort((a,b) => String(b.date || '').localeCompare(String(a.date || '')) || String(b.createdAt || '').localeCompare(String(a.createdAt || '')))
    .slice(0,10);
  return {
    queueVersion:3,
    user:publicAccount_(user),
    accounts:accounts.map(publicAccount_),
    trips:trips.slice((page-1)*pageSize,page*pageSize),
    tripCount:trips.length,
    tripAllowanceTotal:trips.reduce((sum,t) => sum + Number(t.allowance||0),0),
    monthlySummary,
    summaryMonth,
    extraDays:extra.slice((extraPage-1)*pageSize,extraPage*pageSize),
    extraCount:extra.length,
    pendingExtraCount:pendingExtraDays.length,
    pendingExtraIds:pendingExtraDays.map(t => t.id),
    settlements,
    allowance:Number(getSetting_('defaultAllowance',CONFIG.DEFAULT_ALLOWANCE))
  };
}

function ensureSheets_() {
  const ss = getSpreadsheet_();
  const accounts = getSheet_(CONFIG.ACCOUNTS_SHEET,ACCOUNT_HEADERS);
  const trips = getSheet_(CONFIG.TRIPS_SHEET,TRIP_HEADERS);
  const extraDays = getSheet_(CONFIG.EXTRA_DAYS_SHEET,EXTRA_HEADERS);
  getSheet_(CONFIG.SETTLEMENTS_SHEET,SETTLEMENT_HEADERS);
  const settings = getSheet_(CONFIG.SETTINGS_SHEET,SETTING_HEADERS);
  ensureReviewColumns_(trips,TRIP_HEADERS);
  ensureReviewColumns_(extraDays,EXTRA_HEADERS);
  if (accounts.getLastRow() < 2) {
    accounts.appendRow(['admin_primary','علي رجب',CONFIG.PRIMARY_ADMIN_PHONE,'admin','','','',new Date().toISOString()]);
  } else {
    const phones = getAccounts_().map(a => a.phone);
    if (!phones.includes(CONFIG.PRIMARY_ADMIN_PHONE)) {
      accounts.appendRow(['admin_primary','علي رجب',CONFIG.PRIMARY_ADMIN_PHONE,'admin','','','',new Date().toISOString()]);
    }
  }
  if (settings.getLastRow() < 2) settings.appendRow(['defaultAllowance',CONFIG.DEFAULT_ALLOWANCE]);
  migrateJourneysSheet_(ss,trips);
}

function ensureReviewColumns_(sheet,headers) {
  const current = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),1)).getValues()[0].map(value => String(value||'').trim());
  headers.forEach((header) => {
    if (current.indexOf(header) < 0) {
      const column = sheet.getLastColumn() + 1;
      sheet.getRange(1,column).setValue(header).setFontWeight('bold').setBackground('#991b1b').setFontColor('#ffffff');
      current.push(header);
    }
  });
  const statusColumn = current.indexOf('status') + 1;
  if (statusColumn > 0 && sheet.getLastRow() > 1) {
    const statuses = sheet.getRange(2,statusColumn,sheet.getLastRow()-1,1).getValues();
    let changed = false;
    statuses.forEach(row => {
      if (!String(row[0]||'').trim()) {
        row[0] = 'pending';
        changed = true;
      }
    });
    if (changed) sheet.getRange(2,statusColumn,statuses.length,1).setValues(statuses);
  }
}

function normalizeReviewStatus_(value) {
  const status = String(value||'').trim().toLowerCase();
  return status === 'pending' || status === 'rejected' || status === 'approved' ? status : 'pending';
}

function reviewRecord_(type,id,status,actor) {
  const isTrip = type === 'trip';
  const sheet = getSheet_(isTrip ? CONFIG.TRIPS_SHEET : CONFIG.EXTRA_DAYS_SHEET,isTrip ? TRIP_HEADERS : EXTRA_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = rows.findIndex((row,index) => index > 0 && String(row[0]) === String(id));
  if (rowIndex < 1) throw new Error(isTrip ? 'الرحلة غير موجودة.' : 'سجل اليوم الإضافي غير موجود.');
  const headers = rows[0].map(value => String(value||'').trim());
  const statusColumn = headers.indexOf('status') + 1;
  const reviewedByColumn = headers.indexOf('reviewedBy') + 1;
  const reviewedAtColumn = headers.indexOf('reviewedAt') + 1;
  sheet.getRange(rowIndex+1,statusColumn).setValue(status);
  sheet.getRange(rowIndex+1,reviewedByColumn,1,2).setValues([[actor.name,new Date().toISOString()]]);
}

function reviewPendingRecords_(type,ids,status,actor) {
  const isTrip = type === 'trip';
  const sheet = getSheet_(isTrip ? CONFIG.TRIPS_SHEET : CONFIG.EXTRA_DAYS_SHEET,isTrip ? TRIP_HEADERS : EXTRA_HEADERS);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return {updatedCount:0,skippedCount:ids.length,updatedIds:[],updatedAllowance:0};
  const headers = rows[0].map(value => String(value||'').trim());
  const statusIndex = headers.indexOf('status');
  const reviewerIndex = headers.indexOf('reviewedBy');
  const reviewedAtIndex = headers.indexOf('reviewedAt');
  if (statusIndex < 0 || reviewerIndex < 0 || reviewedAtIndex < 0) throw new Error('أعمدة المراجعة غير موجودة في ورقة العمل.');
  const requested = {};
  ids.forEach(id => { requested[String(id)] = true; });
  const now = new Date().toISOString();
  const rowCount = rows.length - 1;
  const statusValues = sheet.getRange(2,statusIndex+1,rowCount,1).getValues();
  const reviewerValues = sheet.getRange(2,reviewerIndex+1,rowCount,1).getValues();
  const reviewedAtValues = sheet.getRange(2,reviewedAtIndex+1,rowCount,1).getValues();
  const allowanceIndex = isTrip ? headers.indexOf('allowance') : -1;
  let updatedCount = 0;
  let updatedAllowance = 0;
  const updatedIds = [];
  for (let i=0;i<rowCount;i++) {
    if (requested[String(rows[i+1][0])] && normalizeReviewStatus_(statusValues[i][0]) === 'pending') {
      statusValues[i][0] = status;
      reviewerValues[i][0] = actor.name;
      reviewedAtValues[i][0] = now;
      updatedCount++;
      updatedIds.push(String(rows[i+1][0]));
      if (allowanceIndex >= 0) updatedAllowance += Number(rows[i+1][allowanceIndex] || 0);
    }
  }
  if (updatedCount) {
    if (reviewerIndex === statusIndex+1 && reviewedAtIndex === reviewerIndex+1) {
      const updates = statusValues.map((row,index) => [row[0],reviewerValues[index][0],reviewedAtValues[index][0]]);
      sheet.getRange(2,statusIndex+1,rowCount,3).setValues(updates);
    } else {
      sheet.getRange(2,statusIndex+1,rowCount,1).setValues(statusValues);
      sheet.getRange(2,reviewerIndex+1,rowCount,1).setValues(reviewerValues);
      sheet.getRange(2,reviewedAtIndex+1,rowCount,1).setValues(reviewedAtValues);
    }
  }
  const uniqueCount = Object.keys(requested).length;
  return {updatedCount,skippedCount:Math.max(0,uniqueCount-updatedCount),updatedIds,updatedAllowance};
}

function migrateJourneysSheet_(ss,trips) {
  const migrationKey = 'journeys_migrated_v1_' + ss.getId();
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(migrationKey) === '1') return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (properties.getProperty(migrationKey) === '1') return;
    if (trips.getLastRow() > 1) {
      properties.setProperty(migrationKey,'1');
      return;
    }
    const old = ss.getSheetByName('Journeys');
    if (!old || old.getLastRow() < 2) {
      properties.setProperty(migrationKey,'1');
      return;
    }
    const accounts = getAccounts_();
    const oldRows = old.getDataRange().getValues().slice(1);
    const converted = oldRows.filter(r => r[0]).map(r => {
      const owner = accounts.find(a => String(a.id) === String(r[1])) || accounts.find(a => a.name === String(r[2]||''));
      return [String(r[0]),owner ? owner.phone:'',String(r[2]||''),String(r[9]||''),formatDate_(r[3]),String(r[5]||''),String(r[4]||''),Number(r[7])||0,String(r[9]||''),'approved','',''];
    });
    if (converted.length) trips.getRange(2,1,converted.length,TRIP_HEADERS.length).setValues(converted);
    properties.setProperty(migrationKey,'1');
  } finally {
    lock.releaseLock();
  }
}

function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('اربط Apps Script بملف Google Sheets أو ضع SPREADSHEET_ID.');
  return ss;
}

function getSheet_(name,headers) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#991b1b').setFontColor('#ffffff');
  }
  sh.setRightToLeft(true);
  if (name === CONFIG.ACCOUNTS_SHEET) sh.getRange('C:C').setNumberFormat('@');
  if (name === CONFIG.TRIPS_SHEET || name === CONFIG.EXTRA_DAYS_SHEET) sh.getRange('B:B').setNumberFormat('@');
  if (name === CONFIG.SETTLEMENTS_SHEET) sh.getRange('B:B').setNumberFormat('@');
  return sh;
}

function getAccounts_() {
  const sh = getSheet_(CONFIG.ACCOUNTS_SHEET,ACCOUNT_HEADERS);
  return sh.getDataRange().getValues().slice(1).filter(row => row[2]).map((row,i) => accountFromRow_(row,i+2));
}

function accountFromRow_(row,rowNumber) {
  return {id:String(row[0]),name:String(row[1]),phone:normalizePhone_(row[2]),role:isPrimaryAdmin_(row[2]) || String(row[3]||'').toLowerCase() === 'admin' ? 'admin' : 'rep',passwordSalt:String(row[4]||''),passwordHash:String(row[5]||''),avatar:String(row[6]||''),createdAt:String(row[7]||''),_row:rowNumber};
}

function findAccountByPhone_(phone) {
  return getAccounts_().find(a => a.phone === normalizePhone_(phone)) || null;
}

function getTrips_() {
  const sh = getSheet_(CONFIG.TRIPS_SHEET,TRIP_HEADERS);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(value => String(value||'').trim());
  const column = name => headers.indexOf(name);
  return values.slice(1).filter(row => row[0]).map(row => ({
    id:String(row[0]),ownerPhone:normalizePhone_(row[1]),repName:String(row[2]||''),timestamp:String(row[3]||''),date:formatDate_(row[4]),place:String(row[5]||''),client:String(row[6]||''),allowance:Number(row[7])||0,createdAt:String(row[8]||''),status:normalizeReviewStatus_(row[column('status')]),reviewedBy:String(row[column('reviewedBy')]||''),reviewedAt:String(row[column('reviewedAt')]||'')
  }));
}

function getExtraDays_() {
  const sh = getSheet_(CONFIG.EXTRA_DAYS_SHEET,EXTRA_HEADERS);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(value => String(value||'').trim());
  const column = name => headers.indexOf(name);
  return values.slice(1).filter(row => row[0]).map(row => ({
    id:String(row[0]),ownerPhone:normalizePhone_(row[1]),repName:String(row[2]||''),timestamp:String(row[3]||''),date:formatDate_(row[4]),reason:String(row[5]||''),proofUrl:String(row[6]||''),createdAt:String(row[7]||''),status:normalizeReviewStatus_(row[column('status')]),reviewedBy:String(row[column('reviewedBy')]||''),reviewedAt:String(row[column('reviewedAt')]||'')
  }));
}

function getSettlements_() {
  const sh = getSheet_(CONFIG.SETTLEMENTS_SHEET,SETTLEMENT_HEADERS);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(value => String(value||'').trim());
  const column = name => headers.indexOf(name);
  return values.slice(1).filter(row => row[0]).map(row => {
    let expenses;
    try {
      expenses = JSON.parse(String(row[column('expensesJson')] || '[]'));
    } catch (error) {
      throw new Error('تعذر قراءة بنود تسوية محفوظة في Google Sheets.');
    }
    if (!Array.isArray(expenses)) throw new Error('تنسيق بنود تسوية محفوظة غير صحيح.');
    return {
      id:String(row[column('id')] || ''),
      ownerPhone:normalizePhone_(row[column('ownerPhone')]),
      repName:String(row[column('repName')] || ''),
      date:formatDate_(row[column('date')]),
      receivedToday:Number(row[column('receivedToday')]) || 0,
      priorBalance:Number(row[column('priorBalance')]) || 0,
      totalCash:Number(row[column('totalCash')]) || 0,
      expenses:expenses.map(item => ({name:String(item.name || ''),amount:Number(item.amount) || 0,gratuity:Number(item.gratuity) || 0})),
      spent:Number(row[column('spent')]) || 0,
      remaining:Number(row[column('remaining')]) || 0,
      createdAt:String(row[column('createdAt')] || '')
    };
  });
}

function pruneSettlements_() {
  const sheet = getSheet_(CONFIG.SETTLEMENTS_SHEET,SETTLEMENT_HEADERS);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 11) return;
  const headers = values[0].map(value => String(value||'').trim());
  const createdAtIndex = headers.indexOf('createdAt');
  const records = values.slice(1)
    .map((row,index) => ({
      rowNumber:index + 2,
      id:String(row[0] || ''),
      createdAt:String(createdAtIndex >= 0 ? row[createdAtIndex] || '' : '')
    }))
    .filter(record => record.id)
    .sort((a,b) => b.createdAt.localeCompare(a.createdAt) || b.id.localeCompare(a.id))
    .slice(0,10);
  const keepRows = {};
  records.forEach(record => { keepRows[record.rowNumber] = true; });
  for (let rowNumber=values.length; rowNumber>=2; rowNumber--) {
    if (!keepRows[rowNumber]) sheet.deleteRow(rowNumber);
  }
}

function saveProofImage_(dataUrl,name,type) {
  if (!dataUrl) return '';
  const match = String(dataUrl).match(/^data:(.*?);base64,(.*)$/);
  if (!match) throw new Error('صيغة صورة الإثبات غير صحيحة.');
  if (String(match[2]).length > 6000000) throw new Error('حجم صورة الإثبات كبير جدًا.');
  const blob = Utilities.newBlob(Utilities.base64Decode(match[2]),type || match[1] || 'image/jpeg',safeFileName_(name || 'extra-proof.jpg'));
  let folder;
  const folders = DriveApp.getFoldersByName('TAGADDOD Extra Days');
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('TAGADDOD Extra Days');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  return file.getUrl();
}

function formatDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value,Session.getScriptTimeZone(),'yyyy-MM-dd');
  return String(value||'').slice(0,10);
}

function publicAccount_(account) {
  return {id:account.id,name:account.name,phone:account.phone,role:isAdmin_(account)?'admin':'rep',avatar:account.avatar||'',firstLogin:!account.passwordHash};
}

function isAdmin_(account) { return Boolean(account && (String(account.role) === 'admin' || isPrimaryAdmin_(account.phone))); }
function isPrimaryAdmin_(phone) { return normalizePhone_(phone) === CONFIG.PRIMARY_ADMIN_PHONE; }
function normalizePhone_(phone) {
  let digits = String(phone||'').replace(/\D/g,'');
  if (digits.length === 10 && /^(10|11|12|15)/.test(digits)) digits = '0' + digits;
  return digits;
}
function requireAdmin_(allowed) { if (!allowed) throw new Error('هذه العملية متاحة للمشرف فقط.'); }
function safeId_(value) { return String(value||'').replace(/[^a-zA-Z0-9_-]/g,'').slice(0,100); }
function safeFileName_(value) { return String(value||'proof.jpg').replace(/[^a-zA-Z0-9_.-]/g,'_').slice(0,120); }

function hashPassword_(password,salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,salt + String(password),Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes);
}
function verifyPassword_(password,salt,expected) { return Boolean(salt && expected && hashPassword_(password,salt) === expected); }

function getSession_(token) {
  if (!token) return null;
  const key = 'SESSION_' + safeId_(token);
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return null;
  try {
    const session = JSON.parse(raw);
    if (!session.expiresAt || session.expiresAt < Date.now()) { PropertiesService.getScriptProperties().deleteProperty(key); return null; }
    return {phone:normalizePhone_(session.phone),key};
  } catch (_) { return null; }
}
function updateSessionPhone_(token,phone) {
  const session = getSession_(token);
  if (!session) throw new Error('انتهت الجلسة.');
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(session.key);
  const data = JSON.parse(raw);
  data.phone = normalizePhone_(phone);
  props.setProperty(session.key,JSON.stringify(data));
}
function getSetting_(key,fallback) {
  const sh = getSheet_(CONFIG.SETTINGS_SHEET,SETTING_HEADERS);
  const row = sh.getDataRange().getValues().find((item,i) => i > 0 && String(item[0]) === key);
  return row ? row[1] : fallback;
}
function setSetting_(key,value) {
  const sh = getSheet_(CONFIG.SETTINGS_SHEET,SETTING_HEADERS);
  const values = sh.getDataRange().getValues();
  const index = values.findIndex((row,i) => i > 0 && String(row[0]) === key);
  if (index > 0) sh.getRange(index+1,2).setValue(value); else sh.appendRow([key,value]);
}
function writeCachedResult_(requestId,result) {
  const cache = CacheService.getScriptCache();
  const key = 'req_' + safeId_(requestId);
  const serialized = JSON.stringify(result);
  const chunkSize = 20000;
  if (serialized.length <= chunkSize) { cache.put(key,serialized,CONFIG.RESULT_CACHE_SECONDS); return; }
  const chunks=[];
  for (let i=0;i<serialized.length;i+=chunkSize) chunks.push(serialized.slice(i,i+chunkSize));
  if (chunks.length > 10) { cache.put(key,JSON.stringify({ok:false,error:'حجم البيانات كبير جدًا للتحميل دفعة واحدة.'}),CONFIG.RESULT_CACHE_SECONDS); return; }
  chunks.forEach((chunk,i) => cache.put(key+'_'+i,chunk,CONFIG.RESULT_CACHE_SECONDS));
  cache.put(key,JSON.stringify({chunkCount:chunks.length}),CONFIG.RESULT_CACHE_SECONDS);
}
function readCachedResult_(requestId) {
  const cache=CacheService.getScriptCache();
  const key='req_'+safeId_(requestId);
  const value=cache.get(key);
  if (!value) return {pending:true};
  const meta=JSON.parse(value);
  if (!meta.chunkCount) return meta;
  let serialized='';
  for (let i=0;i<meta.chunkCount;i++) {
    const chunk=cache.get(key+'_'+i);
    if (chunk===null) return {pending:true};
    serialized += chunk;
  }
  return JSON.parse(serialized);
}
function output_(data,callback) {
  const json=JSON.stringify(data);
  if (callback && /^__delta_cb_[A-Za-z0-9_]+$/.test(callback)) return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}