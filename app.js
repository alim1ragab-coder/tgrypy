/* TAGADDOD Delta Reps - client application. Google Sheets is the source of truth. */
(() => {
  "use strict";

  const CONFIG = {
    appVersion: "1.0.0",
    updateCheckUrl: "https://raw.githubusercontent.com/aliragab-coder/delta-reps/main/version.json",
    sheetUrl: "https://script.google.com/macros/s/AKfycbztDoTUw2zJshF0m46fSnsAH4z8DMunZS1liZktxCzdKgLp5uaF2m9X40L9desEYML_Ig/exec",
    xlsxUrl: "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js",
    sessionKey: "delta_session_token_v5",
    cacheKey: "delta_state_cache_v5",
    queueKey: "delta_mutation_queue_v1",
    pageSize: 10
  };

  const $ = (id) => document.getElementById(id);
  let state = { user: null, accounts: [], trips: [], extraDays: [], settlements: [], monthlySummary: [], summaryMonth: "", tripCount: 0, tripAllowanceTotal: 0, extraCount: 0, pendingExtraCount: 0, pendingExtraIds: [], queueVersion: 0, allowance: 50 };
  let currentPage = 1;
  let extraCurrentPage = 1;
  let loginBusy = false;
  let firstLoginMode = false;
  let lastLoginPhone = "";
  let lastStateError = "";
  let toastTimer = null;
  let actionProgressSequence = 0;
  const activeActionProgress = [];
  let appLoadingDepth = 0;
  let mutationQueue = [];
  let queueSyncBusy = false;
  let activeQueueRequestId = "";
  let xlsxLoadPromise = null;
  let serverMutationChain = Promise.resolve();

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  function queueServerMutation(action) {
    const result = serverMutationChain.then(action, action);
    serverMutationChain = result.then(() => undefined, () => undefined);
    return result;
  }

  const normalizePhone = (value) => {
    let digits = String(value || "").replace(/\D/g, "");
    if (digits.length === 10 && /^(10|11|12|15)/.test(digits)) digits = `0${digits}`;
    return digits;
  };

  const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[ch]));

  const formatTimestamp = (raw) => {
    const fallbackNow = new Date();
    const d = raw ? new Date(raw) : fallbackNow;
    if (Number.isNaN(d.getTime())) return String(raw || "");
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const hours24 = d.getHours();
    const period = hours24 >= 12 ? "م" : "ص";
    const hours12 = String(hours24 % 12 || 12).padStart(2, "0");
    const mins = String(d.getMinutes()).padStart(2, "0");
    const secs = String(d.getSeconds()).padStart(2, "0");
    return `${yyyy}/${mm}/${dd} - ${hours12}:${mins}:${secs} ${period}`;
  };

  const isAdmin = () => state.user?.role === "admin";
  const sessionToken = () => localStorage.getItem(CONFIG.sessionKey) || "";
  window.getActiveUserName = () => state.user?.name || "";
  const reviewLabel = (status) => ({
    pending: "مؤكد في Google Sheets — تحت المراجعة",
    approved: "مقبول",
    rejected: "مرفوض",
    "sync-pending": "ظهر عند المندوب — بانتظار تأكيد الشيت",
    "sync-failed": "تعذر الإرسال—أعد المحاولة"
  }[status] || "تحت المراجعة");

  function createReviewStatus(status) {
    const badge = document.createElement("span");
    badge.className = `review-status review-${status || "pending"}`;
    badge.textContent = reviewLabel(status);
    return badge;
  }

  function showToast(message, kind = "info", duration = 2600) {
    let el = $("appToast");
    if (!el) {
      el = document.createElement("div");
      el.id = "appToast";
      el.className = "app-toast";
      el.setAttribute("role", "status");
      el.setAttribute("aria-live", "polite");
      document.body.appendChild(el);
    }
    clearTimeout(toastTimer);
    el.className = `app-toast ${kind} visible`;
    el.textContent = message;
    if (kind !== "loading") toastTimer = setTimeout(() => el.classList.remove("visible"), duration);
  }
  window.showAppToast = showToast;

  function beginActionProgress(message) {
    const operation = { id: ++actionProgressSequence, message };
    activeActionProgress.push(operation);
    renderActionProgress();
    return () => {
      const index = activeActionProgress.findIndex((item) => item.id === operation.id);
      if (index >= 0) activeActionProgress.splice(index, 1);
      renderActionProgress();
    };
  }
  window.beginAppActionProgress = beginActionProgress;

  function renderActionProgress() {
    const indicator = $("actionProgress");
    if (!indicator) return;
    const current = activeActionProgress[activeActionProgress.length - 1];
    indicator.textContent = current?.message || "";
    indicator.classList.toggle("hidden", !current);
  }

  async function withButtonProgress(button, loadingMessage, action) {
    if (!button || button.dataset.actionPending === "true") return false;
    const originalText = button.textContent;
    const finishProgress = beginActionProgress(loadingMessage);
    button.dataset.actionPending = "true";
    button.setAttribute("aria-busy", "true");
    button.textContent = loadingMessage;
    try {
      return await action(button);
    } catch (error) {
      showToast(error.message || "تعذر تنفيذ العملية.", "error", 6000);
      return false;
    } finally {
      finishProgress();
      delete button.dataset.actionPending;
      button.removeAttribute("aria-busy");
      if (button.isConnected) button.textContent = originalText;
    }
  }

  function bindProgressClick(button, loadingMessage, action) {
    button?.addEventListener("click", (event) => {
      withButtonProgress(event.currentTarget, loadingMessage, action);
    });
  }

  function setAppLoading(enabled) {
    appLoadingDepth = Math.max(0, appLoadingDepth + (enabled ? 1 : -1));
    document.body?.classList.toggle("app-loading", appLoadingDepth > 0);
  }

  function setStatus(message, kind = "info") {
    const el = $("syncStatus");
    if (el) {
      el.textContent = message;
      el.className = `notice ${kind}`;
    }
    showToast(message, kind);
  }

  function setConnectionBadge() {
    const badge = $("connectionStatusBadge");
    if (!badge) return;
    const online = navigator.onLine;
    const status = online ? "متصل بالإنترنت" : "غير متصل";
    const label = badge.querySelector(".connection-status-label");
    if (label) label.textContent = status;
    else badge.textContent = status;
    badge.setAttribute("aria-label", status);
    badge.title = status;
    badge.classList.toggle("status-online", online);
    badge.classList.toggle("status-offline", !online);
  }

  function loadXlsxLibrary() {
    if (typeof XLSX !== "undefined") return Promise.resolve();
    if (!navigator.onLine) return Promise.reject(new Error("اتصل بالإنترنت لتحميل مكتبة Excel."));
    if (xlsxLoadPromise) return xlsxLoadPromise;
    xlsxLoadPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.async = true;
      script.src = CONFIG.xlsxUrl;
      script.onload = () => {
        if (typeof XLSX !== "undefined") return resolve();
        xlsxLoadPromise = null;
        script.remove();
        reject(new Error("تعذر تهيئة مكتبة Excel."));
      };
      script.onerror = () => {
        xlsxLoadPromise = null;
        script.remove();
        reject(new Error("تعذر تحميل مكتبة Excel. تأكد من الاتصال بالإنترنت ثم حاول."));
      };
      document.head.appendChild(script);
    });
    return xlsxLoadPromise;
  }

  function apiGet(params = {}, timeoutMs = 10000) {
    return new Promise((resolve, reject) => {
      if (!CONFIG.sheetUrl) return reject(new Error("لم يتم ضبط رابط Google Apps Script."));
      const callback = `__delta_cb_${Date.now()}_${Math.random().toString(36).slice(2)}`;
      const script = document.createElement("script");
      const query = new URLSearchParams({ ...params, callback });
      let settled = false;
      const finish = (fn, value) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        delete window[callback];
        script.remove();
        fn(value);
      };
      window[callback] = (data) => finish(resolve, data);
      script.onerror = () => finish(reject, new Error("تعذر الاتصال بخدمة Google Sheets."));
      const timer = setTimeout(() => finish(reject, new Error("انتهت مهلة الاتصال بخدمة Google Sheets.")), timeoutMs);
      script.src = `${CONFIG.sheetUrl}?${query.toString()}`;
      document.head.appendChild(script);
    });
  }

  async function apiPost(action, payload = {}, stableRequestId = "") {
    if (!navigator.onLine) throw new Error("الإنترنت فاصل. افتح الإنترنت الأول وبعد كده سجل.");
    if (!CONFIG.sheetUrl) throw new Error("رابط Google Sheets غير مضبوط.");
    const requestId = (stableRequestId || crypto.randomUUID?.() || `req_${Date.now()}_${Math.random().toString(36).slice(2)}`).replace(/[^a-zA-Z0-9_-]/g, "");
    const body = { requestId, action, payload, sessionToken: sessionToken() };
    const controller = new AbortController();
    const sendTimer = setTimeout(() => controller.abort(), 18000);
    try {
      await fetch(CONFIG.sheetUrl, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(body),
        signal: controller.signal
      });
    } catch (error) {
      if (!navigator.onLine) throw new Error("الإنترنت فاصل. افتح الإنترنت الأول وبعد كده سجل.");
      throw new Error("تعذر تأكيد الاتصال بخدمة التسجيل. راجع سجل الرحلات قبل إعادة المحاولة حتى لا يتكرر التسجيل.");
    } finally {
      clearTimeout(sendTimer);
    }
    for (let attempt = 0; attempt < 20; attempt += 1) {
      if (!navigator.onLine) throw new Error("انقطع الإنترنت قبل تأكيد التسجيل. راجع سجل الرحلات قبل إعادة المحاولة.");
      const result = await apiGet({ action: "result", requestId }, 5000);
      if (result && result.pending !== true) {
        if (!result.ok) {
          const error = new Error(result.error || "تعذر تنفيذ العملية.");
          error.code = result.code || "";
          throw error;
        }
        return result;
      }
      await sleep(450);
    }
    throw new Error("التسجيل أُرسل لكن لم يصل تأكيد Google Sheets. راجع سجل الرحلات قبل إعادة المحاولة حتى لا يتكرر التسجيل.");
  }

  function cacheState() {
    try { localStorage.setItem(CONFIG.cacheKey, JSON.stringify(state)); } catch (_) {}
  }

  function clearLegacySavedAnalyses() {
    try {
      localStorage.removeItem("delta_local_lab_analyses_v1");
    } catch (error) {
      showToast("تعذر إزالة التحاليل المحفوظة القديمة من هذا الجهاز.", "warning", 6000);
    }
  }

  function readCachedState() {
    try {
      const saved = JSON.parse(localStorage.getItem(CONFIG.cacheKey) || "null");
      if (saved && typeof saved === "object") {
        delete saved.analyses;
        state = { ...state, ...saved };
      }
    } catch (_) {}
  }

  function readMutationQueue() {
    try {
      const saved = JSON.parse(localStorage.getItem(CONFIG.queueKey) || "[]");
      mutationQueue = Array.isArray(saved) ? saved : [];
      const obsoleteTripReviews = mutationQueue.filter((item) =>
        item.action === "review_records" && item.payload?.type === "trip"
      );
      if (obsoleteTripReviews.length) {
        const remaining = mutationQueue.filter((item) => !obsoleteTripReviews.includes(item));
        if (persistMutationQueue(remaining)) {
          showToast("تم إلغاء قرارات مراجعة الرحلات القديمة؛ الرحلات لم تعد تحتاج موافقة.", "info", 6000);
        }
      }
    } catch (error) {
      mutationQueue = [];
      showToast("تعذر قراءة الطلبات المحفوظة على الجهاز. لا تمسح بيانات التطبيق قبل مراجعة السجل.", "error", 8000);
    }
  }

  function persistMutationQueue(nextQueue) {
    try {
      localStorage.setItem(CONFIG.queueKey, JSON.stringify(nextQueue));
      mutationQueue = nextQueue;
      return true;
    } catch (error) {
      showToast("مساحة حفظ الطلبات على الجهاز غير كافية. أرسل الطلبات المعلّقة أو وفر مساحة ثم حاول.", "error", 8000);
      return false;
    }
  }

  function currentUserQueue() {
    const phone = normalizePhone(state.user?.phone);
    return mutationQueue.filter((item) => normalizePhone(item.userPhone) === phone);
  }

  function queuedReview(type, recordId) {
    return currentUserQueue().find((item) =>
      item.action === "review_records" &&
      item.payload.type === type &&
      item.payload.ids.includes(recordId) &&
      (item.status === "queued" || item.status === "failed")
    ) || null;
  }

  function queuedRecordFor(action, recordId) {
    return currentUserQueue().find((item) =>
      item.action === action &&
      item.localRecordId === recordId &&
      (item.status === "queued" || item.status === "failed")
    ) || null;
  }

  function queuedLocalRecords(action) {
    return currentUserQueue().filter((item) =>
      item.action === action &&
      (item.status === "queued" || item.status === "failed")
    ).map((item) => {
      const payload = item.payload || {};
      const repPhone = normalizePhone(payload.repPhone || item.userPhone);
      const account = state.accounts.find((entry) => normalizePhone(entry.phone) === repPhone);
      return {
        id: item.localRecordId,
        queueRequestId: item.requestId,
        ownerPhone: repPhone,
        repPhone,
        repName: account?.name || (repPhone === normalizePhone(state.user?.phone) ? state.user?.name : ""),
        timestamp: payload.timestamp || item.createdAt,
        date: payload.date,
        place: payload.place,
        client: payload.client,
        allowance: Number(payload.allowance ?? state.allowance),
        reason: payload.reason,
        proofData: payload.proofData || "",
        status: item.status === "failed" ? "sync-failed" : "sync-pending",
        isLocalQueue: true
      };
    });
  }

  function displayedReviewStatus(type, record) {
    const queued = queuedReview(type, record.id);
    if (queued) return queued.status === "failed" ? "sync-failed" : "sync-pending";
    if (record.isLocalQueue) return record.status;
    return record.status || "pending";
  }

  function appendQueueControls(container, item) {
    if (!item) return;
    const button = document.createElement("button");
    button.type = "button";
    button.className = item.status === "failed" ? "btn btn-secondary btn-sm" : "btn btn-danger btn-sm";
    const isSending = item.requestId === activeQueueRequestId;
    button.textContent = isSending
      ? "جاري الإرسال..."
      : item.status === "failed" || item.hasAttempted
        ? "إعادة المحاولة"
        : item.action === "review_records" ? "إلغاء القرار المعلّق" : "إلغاء التسجيل المعلّق";
    button.disabled = isSending;
    button.addEventListener("click", async () => {
      if (button.dataset.actionPending === "true") return;
      if (item.status === "failed" || item.hasAttempted) {
        await withButtonProgress(button, "جاري إعادة الإرسال...", () => retryQueuedMutation(item.requestId));
      } else if (confirm(item.action === "review_records"
        ? "إلغاء هذا القرار سيشمل كل السجلات المرتبطة به في هذه الدفعة. هل تريد المتابعة؟"
        : "هل تريد إلغاء هذا التسجيل المحفوظ على الجهاز؟")) {
        await withButtonProgress(button, "جاري الإلغاء...", () => removeQueuedMutation(item.requestId));
      }
    });
    const controls = document.createElement("div");
    controls.className = "review-actions";
    controls.appendChild(button);
    container.appendChild(controls);
  }

  function newRequestId() {
    return (crypto.randomUUID?.() || `req_${Date.now()}_${Math.random().toString(36).slice(2)}`).replace(/[^a-zA-Z0-9_-]/g, "");
  }

  function enqueueMutation(action, payload) {
    if (!state.user) {
      openLogin();
      return false;
    }
    const requestId = newRequestId();
    const item = {
      requestId,
      action,
      payload,
      userPhone: normalizePhone(state.user.phone),
      createdAt: new Date().toISOString(),
      status: "queued",
      error: ""
    };
    if (action === "add_trip") item.localRecordId = `local_${requestId}`;
    if (action === "add_extra_day") item.localRecordId = `local_${requestId}`;
    if (action === "add_settlement") item.localRecordId = `local_${requestId}`;
    if (!persistMutationQueue([...mutationQueue, item])) return false;
    renderAll();
    const queuedMessage = action === "add_trip"
      ? "تم حفظ الرحلة على الجهاز؛ بانتظار مزامنتها مع Google Sheets."
      : action === "add_settlement"
        ? "تم حفظ التسوية على الجهاز؛ ستُزامَن مع Google Sheets عند عودة الإنترنت."
        : "تم حفظ الطلب؛ حالته معلّقة لحين تأكيده في Google Sheets.";
    showToast(queuedMessage, "warning", 3200);
    if (navigator.onLine) syncMutationQueue().catch((error) => showToast(error.message || "تعذرت مزامنة الطلب المحفوظ.", "error", 7000));
    return true;
  }

  function removeQueuedMutation(requestId) {
    const nextQueue = mutationQueue.filter((item) => item.requestId !== requestId);
    if (!persistMutationQueue(nextQueue)) return false;
    renderAll();
    showToast("تم إلغاء الطلب المعلّق على هذا الجهاز.", "warning");
    return true;
  }

  async function retryQueuedMutation(requestId) {
    const nextQueue = mutationQueue.map((item) => item.requestId === requestId ? { ...item, status: "queued", error: "" } : item);
    if (!persistMutationQueue(nextQueue)) return;
    await syncMutationQueue();
  }

  async function syncMutationQueue() {
    if (queueSyncBusy || !navigator.onLine || !sessionToken() || !state.user) return false;
    const ownerPhone = normalizePhone(state.user.phone);
    const queued = mutationQueue.filter((item) => normalizePhone(item.userPhone) === ownerPhone && item.status === "queued");
    if (!queued.length) return false;
    if (Number(state.queueVersion) !== 3) {
      showToast("الطلبات ما زالت محفوظة على الجهاز. حدّث Code.gs وانشر الإصدار الجديد قبل المزامنة الآمنة.", "error", 8000);
      return false;
    }
    queueSyncBusy = true;
    let didSync = false;
    try {
      while (navigator.onLine) {
        const item = mutationQueue.find((queuedItem) =>
          normalizePhone(queuedItem.userPhone) === ownerPhone &&
          queuedItem.status === "queued"
        );
        if (!item) break;
        const attempting = mutationQueue.map((queuedItem) => queuedItem.requestId === item.requestId
          ? { ...queuedItem, hasAttempted: true }
          : queuedItem);
        if (!persistMutationQueue(attempting)) break;
        activeQueueRequestId = item.requestId;
        renderAll();
        try {
          const result = await queueServerMutation(() => apiPost(item.action, item.payload, item.requestId));
          activeQueueRequestId = "";
          const remaining = mutationQueue.filter((queuedItem) => queuedItem.requestId !== item.requestId);
          if (!persistMutationQueue(remaining)) break;
          didSync = true;
          applyConfirmedReview(item, result);
          if (item.action === "review_records") {
            const verb = item.payload.status === "approved" ? "قبول" : "رفض";
            const updatedCount = Number(result.updatedCount || 0);
            const skippedCount = Number(result.skippedCount || 0);
            showToast(
              skippedCount
                ? `أكد Google Sheets ${verb} ${updatedCount} سجلًا، وتخطّى ${skippedCount} سجلًا لم يعد معلّقًا. حدّث السجل للمراجعة.`
                : `تم تأكيد ${verb} ${updatedCount} سجلًا في Google Sheets.`,
              skippedCount ? "warning" : "success",
              skippedCount ? 7000 : 4200
            );
          } else if (item.action === "add_trip") {
            showToast("تم تسجيل الرحلة في Google Sheets.", "success");
          } else if (item.action === "add_extra_day") {
            showToast("تم تسجيل اليوم الإضافي في Google Sheets وهو الآن تحت مراجعة المشرف.", "success");
          } else if (item.action === "add_settlement") {
            showToast("تم حفظ تسوية العهدة في Google Sheets.", "success");
          }
        } catch (error) {
          const isTransient = !navigator.onLine ||
            /تعذر تأكيد|انتهت مهلة|انقطع الإنترنت|تعذر الاتصال|لم يصل تأكيد|راجع سجل/.test(error.message || "");
          const updated = mutationQueue.map((queuedItem) => queuedItem.requestId === item.requestId ? {
            ...queuedItem,
            status: isTransient ? "queued" : "failed",
            error: error.message || "تعذر تنفيذ الطلب."
          } : queuedItem);
          activeQueueRequestId = "";
          if (!persistMutationQueue(updated)) break;
          renderAll();
          if (!isTransient) showToast(`تعذر إرسال طلب محفوظ: ${error.message}`, "error", 8000);
          if (isTransient) {
            showToast("لم يصل تأكيد Google Sheets بعد. الطلب ما زال محفوظًا وستتم إعادة المحاولة تلقائيًا.", "warning", 7000);
            break;
          }
        }
      }
      if (didSync) {
        await pullState(false);
        renderAll();
      }
      return didSync;
    } finally {
      activeQueueRequestId = "";
      queueSyncBusy = false;
    }
  }

  function applyServerState(data) {
    if (!data || !data.user) throw new Error("استجابة غير مكتملة من الخادم.");
    if (Number(data.queueVersion) !== 3 || !Array.isArray(data.monthlySummary) || !Array.isArray(data.settlements)) {
      throw new Error("نسخة Google Apps Script المنشورة قديمة ولا تدعم تسوية العهدة وملخص الشهر. حدّث Code.gs وانشر إصدارًا جديدًا.");
    }
    const trips = (Array.isArray(data.trips) ? data.trips : [])
      .sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")) || String(b.id || "").localeCompare(String(a.id || "")));
    const totalTripCount = Number(data.tripCount == null ? trips.length : data.tripCount);
    const extraDays = (Array.isArray(data.extraDays) ? data.extraDays : []).map((item) => ({
      ...item,
      status: item.status || "pending"
    }));
    const pendingExtraDays = extraDays.filter((item) => item.status === "pending");
    state = {
      user: data.user,
      accounts: Array.isArray(data.accounts) ? data.accounts : [],
      trips,
      extraDays,
      settlements: Array.isArray(data.settlements) ? data.settlements : [],
      monthlySummary: Array.isArray(data.monthlySummary) ? data.monthlySummary : [],
      summaryMonth: String(data.summaryMonth || ""),
      tripCount: totalTripCount,
      tripAllowanceTotal: Number(data.tripAllowanceTotal || 0),
      extraCount: Number(data.extraCount == null ? extraDays.length : data.extraCount),
      pendingExtraCount: Number(data.pendingExtraCount == null ? pendingExtraDays.length : data.pendingExtraCount),
      pendingExtraIds: Array.isArray(data.pendingExtraIds) ? data.pendingExtraIds : [],
      queueVersion: Number(data.queueVersion) || 0,
      allowance: Number.isFinite(Number(data.allowance)) ? Number(data.allowance) : 50
    };
    cacheState();
    renderAll();
  }

  async function pullState(showMessage = false) {
    lastStateError = "";
    if (!sessionToken()) return false;
    if (!navigator.onLine) {
      lastStateError = "أنت غير متصل بالإنترنت.";
      if (showMessage) setStatus("أنت غير متصل. يتم عرض آخر بيانات محفوظة على هذا الجهاز.", "warning");
      return false;
    }
    try {
      const response = await apiPost("get_state", {
        tripPage: currentPage,
        extraPage: extraCurrentPage,
        summaryMonth: $("summaryMonthInput")?.value || ""
      });
      if (!response?.ok) throw new Error(response?.error || "تعذر تحميل البيانات.");
      applyServerState(response.data);
      if (showMessage) setStatus("تم تحديث البيانات من Google Sheets.", "success");
      return true;
    } catch (error) {
      lastStateError = error.message || "تعذر تحميل بيانات الحساب من Google Sheets.";
      if (showMessage) setStatus(error.message, "error");
      return false;
    }
  }

  function clearInvalidSession() {
    localStorage.removeItem(CONFIG.sessionKey);
    state.user = null;
    state.accounts = [];
    state.trips = [];
    cacheState();
    renderHeader();
  }

  function isInvalidSessionError(message) {
    return message.startsWith("انتهت الجلسة") || message === "الحساب غير موجود.";
  }

  async function restoreSessionInBackground() {
    const loaded = await pullState(false);
    if (loaded) {
      await syncMutationQueue();
      return;
    }

    if (isInvalidSessionError(lastStateError)) {
      clearInvalidSession();
      showToast("انتهت الجلسة. سجل الدخول مرة أخرى.", "warning", 6000);
      return;
    }

    showToast(
      lastStateError
        ? `تعذر تحديث البيانات الآن. ${lastStateError} يتم عرض البيانات المحفوظة إن وُجدت.`
        : "تعذر تحديث البيانات الآن. يتم عرض البيانات المحفوظة إن وُجدت.",
      "warning",
      7000
    );
  }

  function runMutation(action, payload, successMessage) {
    return queueServerMutation(async () => {
      if (!navigator.onLine) {
        showToast("النت فاصل. افتح النت الأول وبعد كده سجل.", "error", 6000);
        return false;
      }
      try {
        await apiPost(action, payload);
        await pullState(false);
        if (successMessage) setStatus(successMessage, "success");
        return true;
      } catch (error) {
        setStatus(error.message, "error");
        return false;
      }
    });
  }

  async function checkForAppUpdate(manual = false) {
    if (!CONFIG.updateCheckUrl) {
      if (manual) showToast("لم يتم تحديد رابط ملف التحديثات بعد.", "warning");
      return;
    }
    if (!navigator.onLine) {
      if (manual) showToast("اتصل بالإنترنت أولاً للبحث عن تحديثات.", "warning");
      return;
    }
    try {
      const sep = CONFIG.updateCheckUrl.includes("?") ? "&" : "?";
      const res = await fetch(`${CONFIG.updateCheckUrl}${sep}t=${Date.now()}`, { cache: "no-store" });
      if (!res.ok) throw new Error("تعذر الوصول لملف التحديثات.");
      const data = await res.json();
      const latestVersion = String(data.version || "").trim();
      const apkUrl = String(data.apkUrl || "").trim();
      const notes = String(data.notes || "").trim();

      if (latestVersion && latestVersion !== CONFIG.appVersion) {
        const msg = `يتوفر إصدار جديد (${latestVersion}) — الإصدار الحالي (${CONFIG.appVersion})${notes ? `\n\nالتحديثات:\n${notes}` : ""}\n\nهل تريد تحميل النسخة الجديدة الآن؟`;
        if (confirm(msg) && apkUrl) {
          window.open(apkUrl, "_blank");
        }
      } else if (manual) {
        showToast(`أنت على أحدث إصدار بالفعل (v${CONFIG.appVersion}).`, "success");
      }
    } catch (_) {
      if (manual) showToast("تعذر التحقق من وجود تحديثات حالياً.", "error");
    }
  }

  function setView(view) {
    ["authSection", "profileModalSection", "appSection"].forEach((id) => $(id)?.classList.add("hidden"));
    if (view === "auth") $("authSection")?.classList.remove("hidden");
    else if (view === "profile") $("profileModalSection")?.classList.remove("hidden");
    else $("appSection")?.classList.remove("hidden");
    renderLoginEntry();
    $("backToSamplesBtn")?.classList.toggle("hidden", view !== "auth");
  }

  function renderLoginEntry() {
    const show = !state.user && !$("appSection")?.classList.contains("hidden");
    $("openLoginFromSamplesBtn")?.classList.toggle("hidden", !show);
    $("guestLoginPrompt")?.classList.toggle("hidden", !show);
  }

  function switchTab(tabId) {
    if (["newTripTab", "extraDayTab", "recordsTab", "settlementTab", "adminTab"].includes(tabId) && !state.user) {
      openLogin();
      return;
    }
    if (tabId === "adminTab" && !isAdmin()) {
      alert("هذه الصفحة متاحة للمشرف فقط.");
      return;
    }
    document.querySelectorAll(".tab-btn").forEach((button) => button.classList.toggle("active", button.dataset.tab === tabId));
    document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.toggle("active", panel.id === tabId));
    if (tabId === "recordsTab") {
      currentPage = 1;
      extraCurrentPage = 1;
      renderTrips();
      renderExtraDays();
      pullState(false);
    }
    if (tabId === "extraDayTab") {
      renderExtraDays();
      pullState(false);
    }
    if (tabId === "settlementTab") {
      renderSettlements();
      pullState(false);
    }
  }

  function renderHeader() {
    const badge = $("userBadge");
    if (badge) {
      const name = badge.querySelector(".user-badge-name");
      const details = badge.querySelector(".user-badge-details");
      const userDetails = state.user ? `${state.user.phone} · ${isAdmin() ? "مشرف" : "مندوب"}` : "";
      if (name && details) {
        name.textContent = state.user?.name || "";
        details.textContent = userDetails;
      } else {
        badge.textContent = state.user ? `${state.user.name} · ${userDetails}` : "زائر";
      }
      if (state.user) {
        const label = `${state.user.name} · ${userDetails}`;
        badge.setAttribute("aria-label", label);
        badge.title = label;
      } else {
        badge.removeAttribute("aria-label");
        badge.removeAttribute("title");
      }
      badge.classList.toggle("hidden", !state.user);
    }
    const avatar = $("headerAvatar");
    if (avatar) {
      avatar.src = state.user?.avatar || "";
      avatar.classList.toggle("hidden", !state.user?.avatar);
    }
    $("logoutBtn")?.classList.toggle("hidden", !state.user);
    $("openChangePassBtn")?.classList.toggle("hidden", !state.user);
    renderLoginEntry();
    document.querySelectorAll(".admin-only").forEach((el) => el.classList.toggle("hidden", !isAdmin()));
    setConnectionBadge();
  }

  function renderRepSelect() {
    const select = $("tripRepSelect");
    const reps = isAdmin() ? state.accounts : [state.user];
    if (select && state.user) {
      select.innerHTML = reps.map((account) => `<option value="${escapeHtml(account.phone)}">${escapeHtml(account.name)}${account.role === "admin" ? " (مشرف)" : ""}</option>`).join("");
      select.disabled = !isAdmin();
    }
    const extraSelect = $("extraDayRepSelect");
    if (extraSelect && state.user) {
      extraSelect.innerHTML = reps.map((account) => `<option value="${escapeHtml(account.phone)}">${escapeHtml(account.name)}${account.role === "admin" ? " (مشرف)" : ""}</option>`).join("");
      extraSelect.disabled = !isAdmin();
    }
    const settlementSelect = $("settlementRepSelect");
    if (settlementSelect && state.user) {
      const settlementReps = isAdmin() ? state.accounts.filter((account) => account.role === "rep") : [state.user];
      settlementSelect.innerHTML = settlementReps
        .map((account) => `<option value="${escapeHtml(account.phone)}">${escapeHtml(account.name)}</option>`)
        .join("");
      settlementSelect.disabled = !isAdmin();
    }
  }

  function renderReps() {
    const list = $("repsList");
    if (!list) return;
    list.innerHTML = "";
    state.accounts.forEach((account) => {
      const li = document.createElement("li");
      const info = document.createElement("div");
      info.className = "rep-row-info";
      if (account.avatar) {
        const img = document.createElement("img");
        img.src = account.avatar;
        img.alt = "";
        img.className = "rep-avatar-mini";
        info.appendChild(img);
      }
      const text = document.createElement("div");
      text.innerHTML = `<strong>${escapeHtml(account.name)}</strong><div class="hint">${escapeHtml(account.phone)} · ${account.role === "admin" ? "مشرف" : "مندوب"}</div>`;
      info.appendChild(text);
      li.appendChild(info);
      if (isAdmin()) {
        const reset = document.createElement("button");
        reset.type = "button";
        reset.className = "btn btn-secondary";
        reset.textContent = "إعادة كلمة السر";
        reset.addEventListener("click", async () => {
          if (reset.dataset.actionPending === "true") return;
          if (!confirm(`سيتم إلغاء كلمة سر ${account.name}. عند الدخول التالي سيتعين عليه تعيين كلمة سر جديدة. هل تريد المتابعة؟`)) return;
          await withButtonProgress(reset, "جاري إعادة ضبط كلمة السر...", () =>
            runMutation("reset_password", { phone: account.phone }, "تمت إعادة ضبط كلمة السر."));
        });
        li.appendChild(reset);
      }
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "btn btn-danger";
      remove.textContent = "حذف";
      const primaryAdmin = account.phone === "01276341407";
      const currentIsPrimary = state.user?.phone === "01276341407";
      remove.disabled = account.phone === state.user?.phone || primaryAdmin || (account.role === "admin" && !currentIsPrimary);
      remove.title = remove.disabled ? "لا يمكن حذف هذا الحساب بالصلاحيات الحالية" : "حذف الحساب";
      remove.addEventListener("click", async () => {
        if (remove.dataset.actionPending === "true") return;
        if (!confirm(`تأكيد حذف حساب ${account.name}؟`)) return;
        await withButtonProgress(remove, "جاري حذف الحساب...", () =>
          runMutation("delete_account", { phone: account.phone }, "تم حذف الحساب بنجاح."));
      });
      li.appendChild(remove);
      list.appendChild(li);
    });
  }

  function getVisibleTrips() {
    const from = isAdmin() ? ($("fromDate")?.value || "") : "";
    const to = isAdmin() ? ($("toDate")?.value || "") : "";
    const query = isAdmin() ? (($("tripSearchInput")?.value || "").trim().toLocaleLowerCase("ar-EG")) : "";
    const currentPhone = normalizePhone(state.user?.phone);
    const currentName = (state.user?.name || "").trim();

    const trips = currentPage === 1 ? [...state.trips, ...queuedLocalRecords("add_trip")] : state.trips;
    return trips.filter((trip) => {
      if (!isAdmin()) {
        const matchPhone = trip.repPhone && normalizePhone(trip.repPhone) === currentPhone;
        const matchOwner = trip.ownerPhone && normalizePhone(trip.ownerPhone) === currentPhone;
        const matchName = currentName && (trip.repName || "").trim() === currentName;
        if (!matchPhone && !matchOwner && !matchName) return false;
      }
      if (from && trip.date < from) return false;
      if (to && trip.date > to) return false;
      const text = `${trip.repName || ""} ${trip.place || ""} ${trip.client || ""}`.toLocaleLowerCase("ar-EG");
      return !query || text.includes(query);
    });
  }

  async function bulkReview(type, status, button) {
    if (!isAdmin()) return showToast("مراجعة السجلات متاحة للمشرف فقط.", "error");
    if (type !== "extra") return showToast("مراجعة بدل السفر ملغاة.", "error");
    const allIds = state.pendingExtraIds;
    const ids = allIds.filter((id) => !queuedReview(type, id));
    if (!ids.length) {
      if (state.pendingExtraCount > 0 && !allIds.length) {
        return showToast("حدّث نشر Code.gs في Apps Script لتفعيل المراجعة الجماعية.", "error", 7000);
      }
      return showToast("لا توجد سجلات معلّقة جديدة لإرسال القرار.", "warning");
    }
    if (ids.length > 5000) return showToast("الدفعة أكبر من الحد المسموح (5000 سجل). حدّث السجل ونفّذ المراجعة على دفعات أصغر.", "error", 7000);
    const noun = "يوم إضافي";
    const verb = status === "approved" ? "قبول" : "رفض";
    if (!confirm(`سيتم ${verb} ${ids.length} ${noun} معلّق دفعة واحدة. سيظل القرار معلّقًا حتى يؤكده Google Sheets. هل تريد المتابعة؟`)) return;
    await withButtonProgress(button, "جاري حفظ القرار...", () =>
      enqueueMutation("review_records", {type, ids, status}));
  }

  async function retryFailedReviews(type, button) {
    const failedIds = new Set(currentUserQueue()
      .filter((item) => item.action === "review_records" && item.payload.type === type && item.status === "failed")
      .map((item) => item.requestId));
    if (!failedIds.size) return;
    const nextQueue = mutationQueue.map((item) => failedIds.has(item.requestId) ? { ...item, status: "queued", error: "" } : item);
    if (!persistMutationQueue(nextQueue)) return;
    await withButtonProgress(button, "جاري إعادة إرسال القرارات...", () => syncMutationQueue());
  }

  function applyConfirmedReview(item, result) {
    if (item.action !== "review_records" || !Array.isArray(result.updatedIds) || !result.updatedIds.length) return;
    const updatedIds = new Set(result.updatedIds.map(String));
    const status = item.payload.status;
    if (item.payload.type !== "extra") return;
    state.extraDays = state.extraDays.map((record) => updatedIds.has(String(record.id)) ? { ...record, status } : record);
    state.pendingExtraIds = state.pendingExtraIds.filter((id) => !updatedIds.has(String(id)));
    state.pendingExtraCount = Math.max(0, state.pendingExtraCount - updatedIds.size);
    cacheState();
    renderAll();
  }

  function formatMoney(value) {
    return `${Number(value || 0).toFixed(2)} جنيه`;
  }

  function renderMonthlySummary() {
    const body = $("monthlySummaryBody");
    const foot = $("monthlySummaryFoot");
    if (!body || !foot) return;
    if (!isAdmin()) {
      body.innerHTML = "";
      foot.innerHTML = "";
      $("monthlySummaryCaption").textContent = "";
      return;
    }
    const month = $("summaryMonthInput")?.value || state.summaryMonth;
    const totals = new Map((state.monthlySummary || []).map((row) => [normalizePhone(row.repPhone), {
      repPhone: normalizePhone(row.repPhone),
      repName: row.repName || "",
      days: Number(row.days || 0),
      allowance: Number(row.allowance || 0)
    }]));
    const localTrips = queuedLocalRecords("add_trip").filter((trip) => !month || String(trip.date || "").slice(0, 7) === month);
    localTrips.forEach((trip) => {
      const phone = normalizePhone(trip.repPhone || trip.ownerPhone);
      const row = totals.get(phone) || { repPhone: phone, repName: trip.repName || "", days: 0, allowance: 0 };
      row.days += 1;
      row.allowance += Number(trip.allowance || 0);
      totals.set(phone, row);
    });
    const rows = [...totals.values()].sort((a, b) => a.repName.localeCompare(b.repName, "ar"));
    body.innerHTML = rows.length
      ? rows.map((row) => `<tr><td data-label="المندوب">${escapeHtml(row.repName)}</td><td data-label="عدد أيام الرحلات">${row.days}</td><td data-label="إجمالي بدل السفر">${escapeHtml(formatMoney(row.allowance))}</td></tr>`).join("")
      : '<tr><td colspan="3">لا توجد رحلات مسجلة لهذا الشهر.</td></tr>';
    const days = rows.reduce((sum, row) => sum + row.days, 0);
    const allowance = rows.reduce((sum, row) => sum + row.allowance, 0);
    foot.innerHTML = rows.length
      ? `<tr><th>الإجمالي</th><th>${days}</th><th>${escapeHtml(formatMoney(allowance))}</th></tr>`
      : "";
    if ($("monthlySummaryCaption")) {
      $("monthlySummaryCaption").textContent = `من أول يوم في ${month || "الشهر المحدد"} إلى آخر يوم فيه${localTrips.length ? ` · ${localTrips.length} رحلة بانتظار المزامنة` : ""}`;
    }
  }

  function renderTrips() {
    const tbody = $("recordsBody");
    if (!tbody) return;
    const rows = getVisibleTrips();
    const localTripCount = queuedLocalRecords("add_trip").length;
    const pages = Math.max(1, Math.ceil((state.tripCount + localTripCount) / CONFIG.pageSize));
    currentPage = Math.min(Math.max(1, currentPage), pages);
    tbody.innerHTML = "";
    rows.forEach((trip) => {
      const tr = document.createElement("tr");
      const formattedTs = formatTimestamp(trip.timestamp);
      const tripCells = [
        { label: "الطابع الزمني", value: formattedTs },
        { label: "التاريخ", value: trip.date || "" },
        { label: "المندوب", value: trip.repName || "" },
        { label: "المكان", value: trip.place || "" },
        { label: "العميل", value: trip.client || "" },
        { label: "بدل السفر", value: `${Number(trip.allowance || 0).toFixed(2)} ج.م` }
      ];
      tr.innerHTML = tripCells.map((cell) => `<td data-label="${cell.label}">${escapeHtml(cell.value)}</td>`).join("");
      const cell = document.createElement("td");
      cell.setAttribute("data-label", "إجراء");
      if (trip.isLocalQueue) {
        appendQueueControls(cell, queuedRecordFor("add_trip", trip.id));
      } else {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "btn btn-danger btn-sm";
        button.textContent = "حذف";
        button.addEventListener("click", async () => {
          if (button.dataset.actionPending === "true") return;
          if (confirm("هل تريد حذف هذه الرحلة نهائيًا من Google Sheets؟")) {
            await withButtonProgress(button, "جاري حذف الرحلة...", () =>
              runMutation("delete_trip", { id: trip.id }, "تم حذف الرحلة بنجاح."));
          }
        });
        cell.appendChild(button);
      }
      tr.appendChild(cell);
      tbody.appendChild(tr);
    });
    const localAllowance = queuedLocalRecords("add_trip").reduce((sum, trip) => sum + Number(trip.allowance || 0), 0);
    $("recordsSummary").textContent = `إجمالي الرحلات: ${state.tripCount} | إجمالي بدل السفر: ${formatMoney(state.tripAllowanceTotal + localAllowance)}${localTripCount ? ` | بانتظار المزامنة: ${localTripCount}` : ""}`;
    $("tripsPageInfo").textContent = `صفحة ${currentPage} من ${pages}`;
    $("prevTripsPageBtn").disabled = currentPage <= 1;
    $("nextTripsPageBtn").disabled = currentPage >= pages;
    $("tripsPaginationBar").classList.toggle("hidden", pages <= 1);
  }

  function renderSettlements() {
    const body = $("settlementsBody");
    if (!body) return;
    const savedRows = (state.settlements || []).map((item) => ({ ...item, localQueueItem: null }));
    const queuedRows = currentUserQueue()
      .filter((item) => item.action === "add_settlement" && (item.status === "queued" || item.status === "failed"))
      .map((item) => {
        const payload = item.payload || {};
        const repPhone = normalizePhone(payload.repPhone || item.userPhone);
        const account = state.accounts.find((entry) => normalizePhone(entry.phone) === repPhone);
        const receivedToday = Number(payload.receivedToday || 0);
        const priorBalance = Number(payload.priorBalance || 0);
        const expenses = Array.isArray(payload.expenses) ? payload.expenses : [];
        const totalCash = receivedToday + priorBalance;
        const spent = expenses.reduce((sum, expense) => sum + Number(expense.amount || 0) + Number(expense.gratuity || 0), 0);
        return {
          id: item.localRecordId,
          date: payload.date,
          repName: account?.name || (repPhone === normalizePhone(state.user?.phone) ? state.user?.name : ""),
          totalCash,
          spent,
          remaining: totalCash - spent,
          localQueueItem: item
        };
      });
    const rows = [...queuedRows, ...savedRows].sort((a, b) =>
      String(b.date || "").localeCompare(String(a.date || "")) ||
      String(b.createdAt || "").localeCompare(String(a.createdAt || ""))
    );
    body.innerHTML = "";
    rows.forEach((settlement) => {
      const tr = document.createElement("tr");
      const values = [
        ["التاريخ", settlement.date || ""],
        ["المندوب", settlement.repName || ""],
        ["إجمالي المبلغ", formatMoney(settlement.totalCash)],
        ["المصروف", formatMoney(settlement.spent)],
        ["المتبقي", formatMoney(settlement.remaining)]
      ];
      tr.innerHTML = values.map(([label, value]) => `<td data-label="${label}">${escapeHtml(value)}</td>`).join("");
      const statusCell = document.createElement("td");
      statusCell.setAttribute("data-label", "الحالة");
      if (settlement.localQueueItem) {
        const status = document.createElement("span");
        status.className = `review-status ${settlement.localQueueItem.status === "failed" ? "review-sync-failed" : "review-sync-pending"}`;
        status.textContent = settlement.localQueueItem.status === "failed" ? "تعذرت المزامنة" : "بانتظار المزامنة";
        statusCell.appendChild(status);
        appendQueueControls(statusCell, settlement.localQueueItem);
      } else {
        statusCell.textContent = "محفوظة في Google Sheets";
      }
      tr.appendChild(statusCell);
      body.appendChild(tr);
    });
    if (!rows.length) body.innerHTML = '<tr><td colspan="6">لا توجد تسويات مسجلة بعد.</td></tr>';
    const pendingCount = queuedRows.length;
    $("settlementsSummary").textContent = `آخر 10 تسويات محفوظة في Google Sheets${pendingCount ? ` · ${pendingCount} بانتظار المزامنة` : ""}`;
  }

  function addSettlementExpenseRow(name = "", amount = "", gratuity = "") {
    const container = $("settlementExpenses");
    if (!container) return;
    const row = document.createElement("div");
    row.className = "settlement-expense-row";
    row.innerHTML = `
      <label>البيان<input type="text" data-expense-name maxlength="150" value="${escapeHtml(name)}" placeholder="مثال: مصنع التونكايا"></label>
      <label>المبلغ (جنيه)<input type="number" data-expense-amount min="0" step="0.01" value="${escapeHtml(amount)}" inputmode="decimal" placeholder="0.00"></label>
      <label>الإكرامية (جنيه)<input type="number" data-expense-gratuity min="0" step="0.01" value="${escapeHtml(gratuity)}" inputmode="decimal" placeholder="0.00"></label>
      <button class="btn btn-danger btn-sm" type="button" data-remove-expense aria-label="حذف البند">حذف البند</button>
    `;
    row.querySelector("[data-remove-expense]").addEventListener("click", () => {
      row.remove();
      updateSettlementTotals();
    });
    container.appendChild(row);
  }

  function updateSettlementTotals() {
    const received = Number($("settlementReceivedToday")?.value || 0);
    const prior = Number($("settlementPriorBalance")?.value || 0);
    const spent = [...document.querySelectorAll(".settlement-expense-row")].reduce((sum, row) => {
      const amount = Number(row.querySelector("[data-expense-amount]").value);
      const gratuity = Number(row.querySelector("[data-expense-gratuity]").value || 0);
      return Number.isFinite(amount) && amount >= 0 && Number.isFinite(gratuity) && gratuity >= 0 ? sum + amount + gratuity : sum;
    }, 0);
    const total = (Number.isFinite(received) && received >= 0 ? received : 0) +
      (Number.isFinite(prior) && prior >= 0 ? prior : 0);
    const remaining = total - spent;
    $("settlementTotalCash").textContent = formatMoney(total);
    $("settlementSpent").textContent = formatMoney(spent);
    $("settlementRemaining").textContent = formatMoney(remaining);
    $("settlementRemaining").classList.toggle("negative", remaining < 0);
  }

  function collectSettlementForm() {
    const date = $("settlementDate").value;
    const repPhone = isAdmin() ? $("settlementRepSelect").value : state.user?.phone;
    const receivedToday = Number($("settlementReceivedToday").value || 0);
    const priorBalance = Number($("settlementPriorBalance").value || 0);
    if (!date) throw new Error("اختر تاريخ التسوية.");
    if (!repPhone) throw new Error("اختر المندوب.");
    if (!Number.isFinite(receivedToday) || receivedToday < 0 || !Number.isFinite(priorBalance) || priorBalance < 0) {
      throw new Error("أدخل مبالغ صحيحة غير سالبة للمبلغ المستلم والرصيد السابق.");
    }
    const expenses = [];
    for (const row of $("settlementExpenses").querySelectorAll(".settlement-expense-row")) {
      const name = row.querySelector("[data-expense-name]").value.trim();
      const amountText = row.querySelector("[data-expense-amount]").value.trim();
      const gratuityText = row.querySelector("[data-expense-gratuity]").value.trim();
      if (!name && !amountText && !gratuityText) continue;
      const amount = Number(amountText);
      const gratuity = Number(gratuityText || 0);
      if (!name || !amountText || !Number.isFinite(amount) || amount < 0 || !Number.isFinite(gratuity) || gratuity < 0) {
        throw new Error("أكمل البيان والمبلغ، وتأكد من صحة قيمة الإكرامية.");
      }
      expenses.push({ name, amount, gratuity });
    }
    if (expenses.length > 100) throw new Error("لا يمكن إضافة أكثر من 100 بند في التسوية.");
    const account = state.accounts.find((item) => normalizePhone(item.phone) === normalizePhone(repPhone));
    const repName = account?.name || state.user?.name || "";
    return { date, repPhone, repName, receivedToday, priorBalance, expenses };
  }

  async function shareSettlementImage() {
    const settlement = collectSettlementForm();
    const totalCash = settlement.receivedToday + settlement.priorBalance;
    const spent = settlement.expenses.reduce((sum, expense) => sum + expense.amount + expense.gratuity, 0);
    const remaining = totalCash - spent;
    const width = 1200;
    const left = 75;
    const right = width - 75;
    const tableWidth = right - left;
    const rowHeights = settlement.expenses.map((expense) => expense.gratuity > 0 ? 110 : 76);
    const tableHeight = 58 + rowHeights.reduce((sum, rowHeight) => sum + rowHeight, 0);
    const height = 660 + tableHeight;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("تعذر تجهيز صورة التسوية.");
    context.fillStyle = "#e5e7eb";
    context.fillRect(0, 0, width, height);
    context.fillStyle = "#ffffff";
    context.fillRect(24, 24, width - 48, height - 48);
    context.direction = "rtl";
    context.textAlign = "center";
    context.fillStyle = "#881337";
    context.font = "800 54px 'Segoe UI', Tahoma, sans-serif";
    context.fillText("تسوية العهدة", width / 2, 105);
    context.fillStyle = "#1f2937";
    context.font = "700 30px 'Segoe UI', Tahoma, sans-serif";
    context.fillText(`المندوب: ${settlement.repName}  |  التاريخ: ${settlement.date}`, width / 2, 165);
    context.fillStyle = "#f3f4f6";
    context.fillRect(70, 205, width - 140, 120);
    context.font = "700 25px 'Segoe UI', Tahoma, sans-serif";
    context.fillStyle = "#374151";
    context.textAlign = "right";
    context.fillText(`المبلغ المستلم اليوم: ${formatMoney(settlement.receivedToday)}`, width - 100, 252);
    context.fillText(`الرصيد المتبقي من رحلات سابقة: ${formatMoney(settlement.priorBalance)}`, width - 100, 296);
    const tableTop = 365;
    const headerHeight = 58;
    const numberDivider = right - 90;
    const amountDivider = left + 240;
    let top = tableTop;
    context.fillStyle = "#d1d5db";
    context.fillRect(left, top, tableWidth, headerHeight);
    context.strokeStyle = "#6b7280";
    context.lineWidth = 2;
    context.strokeRect(left, top, tableWidth, headerHeight);
    context.textAlign = "right";
    context.fillStyle = "#1f2937";
    context.font = "800 28px 'Segoe UI', Tahoma, sans-serif";
    context.fillText("م", right - 28, top + 38);
    context.fillText("البيان", numberDivider - 24, top + 38);
    context.textAlign = "left";
    context.fillText("المبلغ", left + 20, top + 38);
    context.beginPath();
    context.moveTo(numberDivider, top);
    context.lineTo(numberDivider, top + headerHeight);
    context.moveTo(amountDivider, top);
    context.lineTo(amountDivider, top + headerHeight);
    context.stroke();
    top += headerHeight;
    settlement.expenses.forEach((expense, index) => {
      const rowHeight = rowHeights[index];
      context.fillStyle = index % 2 === 0 ? "#ffffff" : "#f3f4f6";
      context.fillRect(left, top, tableWidth, rowHeight);
      context.strokeStyle = "#9ca3af";
      context.lineWidth = 1.5;
      context.strokeRect(left, top, tableWidth, rowHeight);
      context.beginPath();
      context.moveTo(numberDivider, top);
      context.lineTo(numberDivider, top + rowHeight);
      context.moveTo(amountDivider, top);
      context.lineTo(amountDivider, top + rowHeight);
      context.stroke();
      context.textAlign = "right";
      context.fillStyle = "#111827";
      context.font = "700 25px 'Segoe UI', Tahoma, sans-serif";
      context.fillText(String(index + 1), right - 28, top + 39);
      context.fillText(expense.name, numberDivider - 24, top + 39);
      context.textAlign = "left";
      context.fillText(formatMoney(expense.amount), left + 20, top + 39);
      if (expense.gratuity > 0) {
        context.textAlign = "right";
        context.fillStyle = "#6b7280";
        context.font = "600 20px 'Segoe UI', Tahoma, sans-serif";
        context.fillText(`إكرامية ${expense.name}`, numberDivider - 24, top + 82);
        context.textAlign = "left";
        context.fillText(formatMoney(expense.gratuity), left + 20, top + 82);
      }
      top += rowHeight;
    });
    top += 38;
    context.fillStyle = "#f3f4f6";
    context.fillRect(70, top, width - 140, 200);
    context.textAlign = "right";
    context.fillStyle = "#374151";
    context.font = "700 27px 'Segoe UI', Tahoma, sans-serif";
    context.fillText(`إجمالي المبلغ معك: ${formatMoney(totalCash)}`, width - 105, top + 43);
    context.fillText(`إجمالي المصروف اليوم: ${formatMoney(spent)}`, width - 105, top + 96);
    context.fillStyle = remaining < 0 ? "#b91c1c" : "#166534";
    context.font = "800 34px 'Segoe UI', Tahoma, sans-serif";
    context.fillText(`المتبقي معك: ${formatMoney(remaining)}`, width - 105, top + 153);
    context.direction = "ltr";
    context.textAlign = "center";
    context.fillStyle = "#6b7280";
    context.font = "700 20px 'Segoe UI', Tahoma, sans-serif";
    context.fillText("Structured by ALI M. RAGAB  •  TAGADDOD", width / 2, height - 38);
    const dataUrl = canvas.toDataURL("image/png");
    const blob = await new Promise((resolve, reject) => canvas.toBlob((result) =>
      result ? resolve(result) : reject(new Error("تعذر إنشاء ملف صورة للتسوية.")), "image/png"));
    const file = new File([blob], `تسوية_العهدة_${settlement.date}.png`, { type: "image/png" });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({ files: [file], title: "تسوية العهدة" });
        showToast("تمت مشاركة صورة التسوية.", "success");
        return;
      } catch (error) {
        if (error.name === "AbortError") return;
        throw error;
      }
    }
    const link = document.createElement("a");
    link.href = dataUrl;
    link.download = file.name;
    link.click();
    showToast("تم حفظ صورة التسوية على الجهاز.", "success");
  }

  function renderAll() {
    renderHeader();
    renderRepSelect();
    renderReps();
    renderMonthlySummary();
    renderTrips();
    renderExtraDays();
    renderSettlements();
    if ($("tripAllowance")) {
      $("tripAllowance").value = Number(state.allowance).toFixed(2);
      $("tripAllowance").readOnly = true;
      $("tripAllowance").setAttribute("aria-readonly", "true");
    }
    if ($("defaultAllowanceInput")) $("defaultAllowanceInput").value = Number(state.allowance).toFixed(2);
  }

  function syncAuthFields() {
    const phone = normalizePhone($("loginPhoneInput")?.value);
    if (phone !== lastLoginPhone) {
      firstLoginMode = false;
      lastLoginPhone = phone;
      $("setupPasswordInput").value = "";
      $("confirmPasswordInput").value = "";
    }
    const account = state.accounts.find((item) => normalizePhone(item.phone) === phone);
    const accountNeedsSetup = Boolean(account && account.firstLogin);
    $("existingPasswordBox")?.classList.toggle("hidden", firstLoginMode);
    $("firstTimeSetupBox")?.classList.toggle("hidden", !firstLoginMode);
    $("authTitle").textContent = firstLoginMode ? "إنشاء الرقم السري" : accountNeedsSetup ? "تفعيل الحساب لأول مرة" : "تسجيل الدخول";
    $("authHint").textContent = firstLoginMode
      ? "أنشئ رقمًا سريًا من أرقام فقط، 4 أرقام على الأقل."
      : accountNeedsSetup
        ? "اكتب أي رقم سري للمتابعة، وبعدها أنشئ رقمك السري الخاص."
        : "أدخل رقم التليفون والرقم السري المسجل.";
    $("loginPasswordInput").placeholder = accountNeedsSetup && !firstLoginMode ? "اكتب أي رقم للمتابعة" : "الرقم السري";
    $("phoneAuthBtn").textContent = loginBusy
      ? "جاري تسجيل الدخول..."
      : firstLoginMode
        ? "حفظ الرقم السري والدخول"
        : accountNeedsSetup
          ? "متابعة"
          : "تسجيل الدخول";
  }

  function setLoginLoading(isLoading) {
    const section = $("authSection");
    const loginButton = $("phoneAuthBtn");
    section?.classList.toggle("login-in-progress", isLoading);
    section?.setAttribute("aria-busy", String(isLoading));
    $("loginProgressMessage")?.classList.toggle("hidden", !isLoading);
    loginButton?.classList.toggle("login-loading", isLoading);
    section?.querySelectorAll("input, button, select").forEach((control) => {
      if (isLoading) {
        control.dataset.loginPreviousDisabled = String(control.disabled);
        control.disabled = true;
      } else if (control.dataset.loginPreviousDisabled != null) {
        control.disabled = control.dataset.loginPreviousDisabled === "true";
        delete control.dataset.loginPreviousDisabled;
      }
    });
  }

  async function login() {
    if (loginBusy) return;
    if (!navigator.onLine) {
      showToast("النت فاصل. افتح النت الأول وبعد كده سجل الدخول.", "error", 6000);
      return;
    }
    const phone = normalizePhone($("loginPhoneInput").value);
    const password = $("loginPasswordInput").value;
    const newPassword = $("setupPasswordInput").value;
    const confirmPassword = $("confirmPasswordInput").value;
    if (phone.length < 10) return alert("من فضلك أدخل رقم تليفون صحيح.");
    if (!firstLoginMode && !password) return showToast("اكتب أي رقم سري للمتابعة.", "warning");
    if (firstLoginMode && (!/^[0-9]{4,}$/.test(newPassword) || newPassword !== confirmPassword)) {
      return showToast("الرقم السري يجب أن يتكون من 4 أرقام على الأقل، ويجب أن يتطابق مع التأكيد.", "warning", 6000);
    }
    loginBusy = true;
    const finishProgress = beginActionProgress("جاري تسجيل الدخول...");
    setLoginLoading(true);
    try {
      const result = await apiPost("login", { phone, password: firstLoginMode ? "" : password, newPassword: firstLoginMode ? newPassword : "" });
      localStorage.setItem(CONFIG.sessionKey, result.sessionToken);
      $("loginPasswordInput").value = "";
      $("setupPasswordInput").value = "";
      $("confirmPasswordInput").value = "";
      const loaded = await pullState(false);
      if (!loaded) throw new Error(`تم تسجيل الدخول، لكن تعذر تحميل بيانات الحساب: ${lastStateError || "خطأ غير معروف من الخادم."}`);
      setView("app");
      switchTab("analysisTab");
      setStatus("تم تسجيل الدخول ومزامنة البيانات.", "success");
      await syncMutationQueue();
    } catch (error) {
      if (error.code === "FIRST_LOGIN") {
        firstLoginMode = true;
        $("loginPasswordInput").value = "";
        syncAuthFields();
        showToast("أول دخول للحساب. أنشئ رقمًا سريًا من 4 أرقام على الأقل.", "warning", 6500);
      } else if (error.code === "PASSWORD_REQUIRED") {
        firstLoginMode = false;
        syncAuthFields();
        showToast(error.message, "error", 6000);
      } else {
        showToast(error.message, "error", 6000);
      }
    } finally {
      finishProgress();
      loginBusy = false;
      setLoginLoading(false);
      syncAuthFields();
    }
  }

  function openLogin() {
    setView("auth");
    syncAuthFields();
  }

  async function readCompressedAvatar(file) {
    if (!file) return "";
    if (!file.type.startsWith("image/")) throw new Error("الملف المختار ليس صورة.");
    const dataUrl = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
    const image = new Image();
    image.src = dataUrl;
    await image.decode();
    const scale = Math.min(1, 180 / Math.max(image.width, image.height));
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(image.width * scale));
    canvas.height = Math.max(1, Math.round(image.height * scale));
    canvas.getContext("2d").drawImage(image, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL("image/jpeg", 0.55);
  }

  async function saveProfile() {
    if (!navigator.onLine) return showToast("النت فاصل. افتح النت الأول وبعد كده احفظ التعديلات.", "error", 6000);
    try {
      const payload = {
        phone: normalizePhone($("updatePhoneInput").value),
        currentPassword: $("currentPassInput").value,
        newPassword: $("updatedPassInput").value,
        avatar: await readCompressedAvatar($("updateAvatarInput").files?.[0])
      };
      if (!payload.phone && !payload.newPassword && !payload.avatar) return showToast("أدخل تعديلًا واحدًا على الأقل.", "warning");
      if (payload.newPassword && payload.newPassword.length < 4) return showToast("كلمة السر الجديدة يجب أن تكون 4 خانات على الأقل.", "warning");
      const ok = await runMutation("update_profile", payload, "تم تحديث بيانات الحساب بنجاح.");
      if (ok) {
        ["updatePhoneInput", "currentPassInput", "updatedPassInput", "updateAvatarInput"].forEach((id) => { $(id).value = ""; });
        setView("app");
      }
    } catch (error) {
      showToast(error.message || "تعذر تجهيز التعديلات.", "error", 6000);
    }
  }

  function getVisibleExtraDays() {
    const from = isAdmin() ? ($("extraFromDate")?.value || "") : "";
    const to = isAdmin() ? ($("extraToDate")?.value || "") : "";
    const query = isAdmin() ? (($("extraSearchInput")?.value || "").trim().toLocaleLowerCase("ar-EG")) : "";
    const items = extraCurrentPage === 1 ? [...state.extraDays, ...queuedLocalRecords("add_extra_day")] : state.extraDays;
    return items.filter((item) => {
      if (!isAdmin() && normalizePhone(item.ownerPhone) !== normalizePhone(state.user?.phone)) return false;
      if (from && item.date < from) return false;
      if (to && item.date > to) return false;
      const text = `${item.repName || ""} ${item.reason || ""}`.toLocaleLowerCase("ar-EG");
      return !query || text.includes(query);
    });
  }

  function renderExtraDays() {
    const tbody = $("extraRecordsBody");
    if (!tbody) return;
    tbody.innerHTML = "";
    const visibleItems = getVisibleExtraDays();
    visibleItems.forEach((item) => {
      const tr = document.createElement("tr");
      [["الطابع الزمني",formatTimestamp(item.timestamp)],["التاريخ",item.date||""],["المندوب",item.repName||""],["السبب / الملاحظات",item.reason||""]].forEach((cell) => {
        const td = document.createElement("td");
        td.setAttribute("data-label",cell[0]);
        td.textContent = cell[1];
        tr.appendChild(td);
      });
      const proofCell = document.createElement("td");
      proofCell.setAttribute("data-label","صورة الإثبات");
      if (item.isLocalQueue && item.proofData) {
        proofCell.textContent = "مرفق محفوظ على الجهاز";
      } else if (item.proofUrl) {
        const link = document.createElement("a");
        link.className = "proof-link";
        link.href = item.proofUrl;
        link.target = "_blank";
        link.textContent = "عرض الصورة";
        proofCell.appendChild(link);
      } else {
        proofCell.textContent = "لا توجد صورة";
      }
      tr.appendChild(proofCell);
      const statusCell = document.createElement("td");
      statusCell.setAttribute("data-label","الحالة");
      const reviewQueueItem = !item.isLocalQueue ? queuedReview("extra", item.id) : null;
      const itemStatus = displayedReviewStatus("extra", item);
      if (isAdmin() || itemStatus === "approved" || itemStatus === "rejected") {
        statusCell.appendChild(createReviewStatus(itemStatus));
        if (reviewQueueItem) appendQueueControls(statusCell, reviewQueueItem);
      }
      if (isAdmin() && item.status === "pending" && !reviewQueueItem) {
        const controls = document.createElement("div");
        controls.className = "review-actions";
        [["approved","قبول"],["rejected","رفض"]].forEach(([status,label]) => {
          const button = document.createElement("button");
          button.type = "button";
          button.className = status === "approved" ? "btn btn-success btn-sm" : "btn btn-danger btn-sm";
          button.textContent = label;
          button.addEventListener("click", async () => {
            if (button.dataset.actionPending === "true") return;
            if (!confirm(status === "approved" ? "هل تريد قبول اليوم الإضافي؟" : "هل تريد رفض اليوم الإضافي؟")) return;
            await withButtonProgress(button, "جاري حفظ القرار...", () =>
              enqueueMutation("review_records", {type:"extra",ids:[item.id],status}));
          });
          controls.appendChild(button);
        });
        statusCell.appendChild(controls);
      }
      tr.appendChild(statusCell);
      if (isAdmin() || normalizePhone(item.ownerPhone) === normalizePhone(state.user?.phone)) {
        const cell = document.createElement("td");
        cell.setAttribute("data-label", "إجراء");
        if (item.isLocalQueue) {
          appendQueueControls(cell, queuedRecordFor("add_extra_day", item.id));
        } else {
          const button = document.createElement("button");
          button.type = "button";
          button.className = "btn btn-danger btn-sm";
          button.textContent = "حذف";
          button.disabled = Boolean(reviewQueueItem);
          button.title = reviewQueueItem ? "ألغِ قرار المراجعة المعلّق أولًا" : "حذف سجل اليوم الإضافي";
          button.addEventListener("click", async () => {
            if (button.dataset.actionPending === "true") return;
            if (confirm("هل تريد حذف سجل اليوم الإضافي؟")) {
              await withButtonProgress(button, "جاري الحذف...", () =>
                runMutation("delete_extra_day", { id: item.id }, "تم حذف اليوم الإضافي بنجاح."));
            }
          });
          cell.appendChild(button);
        }
        tr.appendChild(cell);
      }
      tbody.appendChild(tr);
    });
    const localExtraCount = queuedLocalRecords("add_extra_day").length;
    const pages = Math.max(1, Math.ceil((state.extraCount + localExtraCount) / CONFIG.pageSize));
    extraCurrentPage = Math.min(Math.max(1, extraCurrentPage), pages);
    const queuedReviewCount = state.pendingExtraIds.filter((id) => queuedReview("extra", id)).length;
    $("extraRecordsSummary").textContent = `إجمالي الأيام الإضافية: ${state.extraCount} | تحت المراجعة: ${state.pendingExtraCount}${localExtraCount ? ` | محفوظة على الجهاز: ${localExtraCount}` : ""}${queuedReviewCount ? ` | قرارات معلّقة للمزامنة: ${queuedReviewCount}` : ""}`;
    $("extraBulkReviewControls")?.classList.toggle("hidden", !isAdmin() || !(state.pendingExtraIds.length || state.pendingExtraCount));
    if ($("extraPendingReviewCount")) $("extraPendingReviewCount").textContent = `السجلات المعلّقة على الخادم: ${state.pendingExtraIds.length || state.pendingExtraCount}`;
    $("retryFailedExtraBtn")?.classList.toggle("hidden", !currentUserQueue().some((item) => item.action === "review_records" && item.payload.type === "extra" && item.status === "failed"));
    $("extraPageInfo").textContent = `صفحة ${extraCurrentPage} من ${pages}`;
    $("prevExtraPageBtn").disabled = extraCurrentPage <= 1;
    $("nextExtraPageBtn").disabled = extraCurrentPage >= pages;
  }

  async function exportExtraExcel() {
    if (!isAdmin()) return showToast("تصدير Excel متاح للمشرف فقط.", "error");
    const from = $("extraFromDate").value;
    const to = $("extraToDate").value;
    if (!from || !to) return showToast("حدد تاريخ البداية والنهاية.", "warning");
    try {
      await loadXlsxLibrary();
      const result = await apiPost("list_extra_days", { from, to });
      if (!result?.ok) throw new Error(result?.error || "تعذر تحميل السجلات.");
      const aoa = [
      ["تقرير الأيام الإضافية - TAGADDOD", "", "", "", "", ""],
      ["الطابع الزمني","التاريخ","المندوب","السبب / الملاحظات","صورة الإثبات","الحالة"],
      ...result.extraDays.map((item) => [formatTimestamp(item.timestamp),item.date,item.repName,item.reason,item.proofUrl,reviewLabel(item.status)])
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [{s:{r:0,c:0},e:{r:0,c:5}}];
    ws["!cols"] = [{wch:24},{wch:15},{wch:22},{wch:35},{wch:45},{wch:18}];
    ws["!views"] = [{RTL:true}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"الأيام الإضافية");
      XLSX.writeFile(wb,`الأيام_الإضافية_${from}_${to}.xlsx`);
    showToast("تم تصدير سجل الأيام الإضافية بنجاح.", "success");
    } catch (error) {
      showToast(error.message || "تعذر تصدير الأيام الإضافية.", "error", 6000);
    }
  }

  async function exportExcel() {
    if (!isAdmin()) return showToast("تصدير Excel متاح للمشرف فقط.", "error");
    const from = $("fromDate").value;
    const to = $("toDate").value;
    if (!from || !to) return showToast("حدد تاريخ البداية والنهاية قبل التصدير.", "warning");
    if (from > to) return showToast("تاريخ البداية يجب ألا يكون بعد تاريخ النهاية.", "warning");

    try {
      await loadXlsxLibrary();
      const listResult = await apiPost("list_trips", { from, to });
      if (!listResult?.ok) throw new Error(listResult?.error || "تعذر تحميل الرحلات للتصدير.");
      const rows = Array.isArray(listResult.trips) ? listResult.trips : [];
      const repNames = [...new Set(rows.map((trip) => trip.repName).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ar"));
      const totalAllowance = rows.reduce((sum, trip) => sum + Number(trip.allowance || 0), 0);
      const aoa = [
        ["تقرير رحلات مشغلي الحركة - TAGADDOD", "", "", "", "", ""],
        ["الطابع الزمني", "تاريخ الرحلة", "اسم المندوب", "المكان / المحافظة", "اسم العميل", "بدل السفر (ج.م)"],
        ...rows.map((trip) => [formatTimestamp(trip.timestamp), trip.date || "", trip.repName || "", trip.place || "", trip.client || "", Number(trip.allowance || 0)])
      ];
      const summaryTitleRow = aoa.length + 2;
      aoa.push([], ["ملخص بدل السفر لكل الرحلات", "", ""]);
      const summaryHeaderRow = aoa.length + 1;
      aoa.push(["اسم المندوب", "عدد الرحلات", "إجمالي بدل السفر (ج.م)"]);
      repNames.forEach((name) => {
        const own = rows.filter((trip) => trip.repName === name);
        aoa.push([name, own.length, own.reduce((sum, trip) => sum + Number(trip.allowance || 0), 0)]);
      });
      const grandTotalRow = aoa.length + 1;
      aoa.push(["الإجمالي العام", rows.length, totalAllowance]);

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws["!views"] = [{ RTL: true }];
      ws["!cols"] = [{ wch: 24 }, { wch: 15 }, { wch: 23 }, { wch: 23 }, { wch: 27 }, { wch: 19 }];
      ws["!rows"] = aoa.map((_, index) => ({ hpt: index === 0 || index + 1 === summaryTitleRow || index + 1 === summaryHeaderRow ? 26 : 22 }));
      ws["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
        { s: { r: summaryTitleRow - 1, c: 0 }, e: { r: summaryTitleRow - 1, c: 2 } }
      ];

      const border = {
        top: { style: "thin", color: { rgb: "9CA3AF" } },
        bottom: { style: "thin", color: { rgb: "9CA3AF" } },
        left: { style: "thin", color: { rgb: "9CA3AF" } },
        right: { style: "thin", color: { rgb: "9CA3AF" } }
      };
      const setStyle = (r, c, style) => {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (cell) cell.s = { ...style, border, alignment: { vertical: "center", horizontal: "center", wrapText: true, ...(style.alignment || {}) } };
      };
      const titleStyle = { font: { name: "Arial", sz: 15, bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "991B1B" } } };
      const headerStyle = { font: { name: "Arial", sz: 11, bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "1F2937" } } };
      const totalStyle = { font: { name: "Arial", sz: 11, bold: true, color: { rgb: "991B1B" } }, fill: { fgColor: { rgb: "FEE2E2" } } };
      const bodyStyle = { font: { name: "Arial", sz: 10, color: { rgb: "111827" } }, fill: { fgColor: { rgb: "FFFFFF" } } };
      for (let c = 0; c < 6; c += 1) setStyle(0, c, titleStyle);
      for (let c = 0; c < 6; c += 1) setStyle(1, c, headerStyle);
      for (let r = 2; r < rows.length + 2; r += 1) {
        for (let c = 0; c < 6; c += 1) {
          const style = { ...bodyStyle, fill: { fgColor: { rgb: r % 2 === 0 ? "F3F4F6" : "FFFFFF" } } };
          if (c === 5) style.numFmt = "#,##0.00";
          setStyle(r, c, style);
        }
      }
      for (let c = 0; c < 3; c += 1) setStyle(summaryTitleRow - 1, c, titleStyle);
      for (let c = 0; c < 3; c += 1) setStyle(summaryHeaderRow - 1, c, headerStyle);
      const summaryStart = summaryHeaderRow;
      for (let r = summaryStart; r < grandTotalRow - 1; r += 1) {
        for (let c = 0; c < 3; c += 1) {
          const style = { ...bodyStyle, fill: { fgColor: { rgb: r % 2 === 0 ? "F3F4F6" : "FFFFFF" } } };
          if (c === 2) style.numFmt = "#,##0.00";
          setStyle(r, c, style);
        }
      }
      for (let c = 0; c < 3; c += 1) {
        const style = { ...totalStyle };
        if (c === 2) style.numFmt = "#,##0.00";
        setStyle(grandTotalRow - 1, c, style);
      }
      ws["!autofilter"] = { ref: `A2:F${Math.max(2, rows.length + 2)}` };
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "تقرير الرحلات");
      XLSX.writeFile(wb, `تقرير_رحلات_المناديب_${from}_${to}.xlsx`);
      showToast(`تم تصدير ${rows.length} رحلة وإجمالي بدل السفر لكل الرحلات.`, "success");
    } catch (error) {
      showToast(`تعذر تصدير الملف: ${error.message || "خطأ غير معروف"}`, "error", 6000);
    }
  }

  function readProofImage(file) {
    if (!file) return Promise.resolve({ dataUrl: "", name: "", type: "" });
    if (!file.type.startsWith("image/")) return Promise.reject(new Error("صورة الإثبات يجب أن تكون ملف صورة."));
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async () => {
        try {
          const image = new Image();
          image.src = String(reader.result);
          await image.decode();
          const scale = Math.min(1, 1280 / Math.max(image.width, image.height));
          const canvas = document.createElement("canvas");
          canvas.width = Math.max(1, Math.round(image.width * scale));
          canvas.height = Math.max(1, Math.round(image.height * scale));
          const context = canvas.getContext("2d");
          if (!context) throw new Error("تعذر تجهيز صورة الإثبات.");
          context.fillStyle = "#ffffff";
          context.fillRect(0, 0, canvas.width, canvas.height);
          context.drawImage(image, 0, 0, canvas.width, canvas.height);
          let quality = 0.72;
          let dataUrl = canvas.toDataURL("image/jpeg", quality);
          while (dataUrl.length > 600000 && quality > 0.42) {
            quality -= 0.08;
            dataUrl = canvas.toDataURL("image/jpeg", quality);
          }
          if (dataUrl.length > 600000) throw new Error("حجم صورة الإثبات كبير. اختر صورة أصغر ثم حاول مرة أخرى.");
          resolve({ dataUrl, name: `${file.name.replace(/\.[^.]+$/, "")}.jpg`, type: "image/jpeg" });
        } catch (error) {
          reject(new Error(error.message || "تعذر تجهيز صورة الإثبات."));
        }
      };
      reader.onerror = () => reject(new Error("تعذر قراءة صورة الإثبات."));
      reader.readAsDataURL(file);
    });
  }

  function bindEvents() {
    bindProgressClick($("checkUpdateBtn"), "جاري الفحص...", () => checkForAppUpdate(true));
    $("loginPhoneInput").addEventListener("input", syncAuthFields);
    $("loginPhoneInput").addEventListener("change", syncAuthFields);
    $("loginForm").addEventListener("submit", (event) => { event.preventDefault(); login(); });
    $("openLoginFromSamplesBtn").addEventListener("click", openLogin);
    $("guestLoginBtn").addEventListener("click", openLogin);
    ["backToSamplesBtn", "cancelLoginBtn"].forEach((id) => $(id).addEventListener("click", () => { setView("app"); switchTab("analysisTab"); }));
    $("openChangePassBtn").addEventListener("click", () => { $("updatePhoneInput").value = ""; $("updatePhoneInput").placeholder = state.user?.phone || "اتركه فارغًا دون تغيير"; setView("profile"); });
    $("closeProfileModalBtn").addEventListener("click", () => setView("app"));
    bindProgressClick($("saveProfileChangesBtn"), "جاري تجهيز التعديلات...", saveProfile);
    $("logoutBtn").addEventListener("click", () => {
      if (navigator.onLine && sessionToken()) apiPost("logout", {}).catch(() => {});
      localStorage.removeItem(CONFIG.sessionKey);
      firstLoginMode = false;
      state.user = null;
      state.accounts = [];
      state.trips = [];
      cacheState();
      renderHeader();
      setView("app");
      switchTab("analysisTab");
      showToast("تم تسجيل الخروج.", "success");
    });
    document.querySelectorAll(".tab-btn").forEach((button) => button.addEventListener("click", () => switchTab(button.dataset.tab)));
    $("tripForm").addEventListener("submit", (event) => {
      event.preventDefault();
      const form = event.currentTarget;
      withButtonProgress(event.submitter || form.querySelector('button[type="submit"]'), "جاري حفظ الرحلة...", async () => {
        if (!state.user) return openLogin();
        const date = $("tripDate").value;
        const place = $("tripPlace").value.trim();
        const client = $("tripClient").value.trim();
        const allowance = Number($("tripAllowance").value);
        if (!date || !place || !client || !Number.isFinite(allowance) || allowance < 0) return showToast("راجع بيانات الرحلة والبدل قبل التسجيل.", "warning");
        const repPhone = isAdmin() ? $("tripRepSelect").value : state.user.phone;
        const ok = await enqueueMutation("add_trip", { date, place, client, allowance, repPhone, timestamp: formatTimestamp(new Date()) });
        if (ok) { $("tripPlace").value = ""; $("tripClient").value = ""; currentPage = 1; renderTrips(); }
      });
    });
    $("extraDayForm").addEventListener("submit", (event) => {
      event.preventDefault();
      const form = event.currentTarget;
      withButtonProgress(event.submitter || form.querySelector('button[type="submit"]'), "جاري تجهيز اليوم الإضافي...", async () => {
        if (!state.user) return openLogin();
        const date = $("extraDayDate").value;
        const reason = $("extraDayReason").value.trim();
        if (!date || !reason) return showToast("أدخل تاريخ اليوم وسبب التكليف.", "warning");
        try {
          const proof = await readProofImage($("extraDayProofInput").files?.[0]);
          const repPhone = isAdmin() ? $("extraDayRepSelect").value : state.user.phone;
          const ok = await enqueueMutation("add_extra_day", { date, reason, repPhone, proofData: proof.dataUrl, proofName: proof.name, proofType: proof.type });
          if (ok) { $("extraDayReason").value = ""; $("extraDayProofInput").value = ""; extraCurrentPage = 1; renderExtraDays(); }
        } catch (error) {
          showToast(error.message || "تعذر تجهيز صورة الإثبات.", "error", 6000);
        }
      });
    });
    $("settlementForm").addEventListener("submit", (event) => {
      event.preventDefault();
      const form = event.currentTarget;
      withButtonProgress($("saveSettlementBtn") || form.querySelector('button[type="submit"]'), "جاري حفظ التسوية...", async () => {
        if (!state.user) return openLogin();
        try {
          const settlement = collectSettlementForm();
          const ok = await enqueueMutation("add_settlement", settlement);
          if (ok) {
            $("settlementReceivedToday").value = "0";
            $("settlementPriorBalance").value = "0";
            $("settlementExpenses").innerHTML = "";
            addSettlementExpenseRow();
            updateSettlementTotals();
            renderSettlements();
          }
        } catch (error) {
          showToast(error.message || "راجع بيانات التسوية.", "warning", 6000);
        }
      });
    });
    $("settlementForm").addEventListener("input", updateSettlementTotals);
    $("addSettlementExpenseBtn").addEventListener("click", () => addSettlementExpenseRow());
    bindProgressClick($("shareSettlementBtn"), "جاري تجهيز صورة التسوية...", shareSettlementImage);
    $("addRepForm").addEventListener("submit", (event) => {
      event.preventDefault();
      if (!isAdmin()) return alert("غير مصرح.");
      const form = event.currentTarget;
      withButtonProgress(event.submitter || form.querySelector('button[type="submit"]'), "جاري إضافة الحساب...", async () => {
        const name = $("newRepName").value.trim();
        const phone = normalizePhone($("newRepPhone").value);
        const role = $("newRepRole").value;
        if (!name || phone.length < 10) return showToast("أدخل اسمًا ورقم تليفون صحيحًا.", "warning");
        const ok = await runMutation("add_account", { name, phone, role }, "تم إضافة الحساب بنجاح.");
        if (ok) { $("newRepName").value = ""; $("newRepPhone").value = ""; }
      });
    });
    bindProgressClick($("saveAllowanceBtn"), "جاري حفظ بدل السفر...", async () => {
      const allowance = Number($("defaultAllowanceInput").value);
      if (!Number.isFinite(allowance) || allowance < 0) return showToast("أدخل قيمة بدل صحيحة.", "warning");
      await runMutation("set_allowance", { allowance }, "تم حفظ بدل السفر الافتراضي بنجاح.");
    });
    ["fromDate", "toDate"].forEach((id) => $(id).addEventListener("change", () => { currentPage = 1; renderTrips(); }));
    $("tripSearchInput").addEventListener("input", () => { currentPage = 1; renderTrips(); });
    bindProgressClick($("prevTripsPageBtn"), "جاري تحديث الرحلات...", async () => { currentPage -= 1; await pullState(true); });
    bindProgressClick($("nextTripsPageBtn"), "جاري تحديث الرحلات...", async () => { currentPage += 1; await pullState(true); });
    bindProgressClick($("exportExcelBtn"), "جاري تجهيز ملف Excel...", exportExcel);
    bindProgressClick($("syncNowBtn"), "جاري تحديث البيانات...", async () => { await pullState(true); await syncMutationQueue(); });
    ["extraFromDate","extraToDate"].forEach((id) => $(id).addEventListener("change", () => { extraCurrentPage = 1; renderExtraDays(); }));
    $("extraSearchInput").addEventListener("input", () => { extraCurrentPage = 1; renderExtraDays(); });
    bindProgressClick($("prevExtraPageBtn"), "جاري تحديث الأيام الإضافية...", async () => { extraCurrentPage -= 1; await pullState(true); });
    bindProgressClick($("nextExtraPageBtn"), "جاري تحديث الأيام الإضافية...", async () => { extraCurrentPage += 1; await pullState(true); });
    bindProgressClick($("exportExtraExcelBtn"), "جاري تجهيز ملف Excel...", exportExtraExcel);
    bindProgressClick($("syncExtraNowBtn"), "جاري تحديث البيانات...", async () => { await pullState(true); await syncMutationQueue(); });
    $("approveAllExtraBtn").addEventListener("click", (event) => bulkReview("extra", "approved", event.currentTarget));
    $("rejectAllExtraBtn").addEventListener("click", (event) => bulkReview("extra", "rejected", event.currentTarget));
    $("retryFailedExtraBtn").addEventListener("click", (event) => retryFailedReviews("extra", event.currentTarget));
    const today = new Date();
    const todayDate = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
    $("tripDate").value = todayDate;
    $("extraDayDate").value = todayDate;
    $("settlementDate").value = todayDate;
    $("summaryMonthInput").value = todayDate.slice(0, 7);
    addSettlementExpenseRow();
    $("summaryMonthInput").addEventListener("change", async () => {
      renderMonthlySummary();
      await pullState(false);
    });
    window.addEventListener("online", async () => {
      setConnectionBadge();
      await pullState(true);
      await syncMutationQueue();
    });
    window.addEventListener("offline", () => {
      setConnectionBadge();
      showToast("الاتصال انقطع. الرحلات والتسويات الجديدة ستُحفظ على الجهاز وتُزامَن عند عودة الإنترنت.", "warning", 7000);
    });
    window.setInterval(() => { if (navigator.onLine) syncMutationQueue(); }, 15000);
  }

  function init() {
    clearLegacySavedAnalyses();
    readCachedState();
    readMutationQueue();
    bindEvents();
    renderHeader();
    renderAll();
    setView("app");
    switchTab("analysisTab");
    setConnectionBadge();

    if (sessionToken()) {
      if (navigator.onLine) {
        restoreSessionInBackground().catch((error) => {
          showToast(error.message || "تعذر تحديث البيانات الآن.", "error", 7000);
        });
      } else {
        setStatus("غير متصل. تظهر البيانات المحفوظة والطلبات المعلّقة حتى عودة الاتصال.", "warning");
      }
    } else {
      state.user = null;
      state.accounts = [];
      state.trips = [];
      renderHeader();
    }
    checkForAppUpdate(false);
  }

  document.addEventListener("DOMContentLoaded", init);
})();