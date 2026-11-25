/* TAGADDOD Delta Reps - local oil sample analysis & share image */
(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  let hasCalculated = false;

  async function withAnalysisButtonProgress(button, loadingMessage, action) {
    if (!button || button.dataset.actionPending === "true") return false;
    const originalText = button.textContent;
    const finishProgress = window.beginAppActionProgress?.(loadingMessage) || (() => {});
    button.dataset.actionPending = "true";
    button.setAttribute("aria-busy", "true");
    button.textContent = loadingMessage;
    try {
      return await action();
    } catch (error) {
      window.showAppToast?.(error.message || "تعذر تنفيذ العملية.", "error", 6000);
      return false;
    } finally {
      finishProgress();
      delete button.dataset.actionPending;
      button.removeAttribute("aria-busy");
      if (button.isConnected) button.textContent = originalText;
    }
  }

  const num = (id) => {
    const v = Number($(id)?.value);
    return Number.isFinite(v) ? v : 0;
  };

  const cleanFloat = (v) => Math.round((Number(v || 0) + Number.EPSILON) * 1e9) / 1e9;

  function computeAnalysis() {
    const clientName = ($("labClientName")?.value || "").trim() || "عميل بدون اسم";
    const acidsInput = num("labAcids");
    const emptyWeight = num("labEmptyWeight");
    const beforeBoil = num("labBeforeBoil");
    const afterBoil = num("labAfterBoil");
    const sedimentWeight = num("labSedimentWeight");

    const sampleDenom = cleanFloat(beforeBoil - emptyWeight);
    const waterDiff = cleanFloat(beforeBoil - afterBoil);
    const sedimentDiff = cleanFloat(sedimentWeight - emptyWeight);

    const waterPct = sampleDenom !== 0 ? Math.max(0, cleanFloat((waterDiff / sampleDenom) * 100)) : 0;
    const sedimentPct = sampleDenom !== 0 ? Math.max(0, cleanFloat((sedimentDiff / sampleDenom) * 100)) : 0;
    const acidPct = Math.max(0, cleanFloat(acidsInput));

    const acidPenalty = acidPct > 3 ? cleanFloat((acidPct - 3) * 10) : 0;
    const wsSum = cleanFloat(waterPct + sedimentPct);
    const wsPenalty = wsSum > 2 ? cleanFloat((wsSum - 2) * 10) : 0;
    const deductionPerTon = Math.round(cleanFloat(acidPenalty + wsPenalty));

    return {
      clientName,
      emptyWeight,
      beforeBoil,
      afterBoil,
      sedimentWeight,
      waterPct,
      sedimentPct,
      acidPct,
      deductionPerTon
    };
  }

  function updateDisplay() {
    if (!hasCalculated) {
      if ($("resWaterPct")) $("resWaterPct").textContent = "0.00%";
      if ($("resSedimentPct")) $("resSedimentPct").textContent = "0.00%";
      if ($("resAcidPct")) $("resAcidPct").textContent = "0.00%";
      if ($("resDeductionPerTon")) $("resDeductionPerTon").textContent = "0";
      return;
    }
    const res = computeAnalysis();
    if ($("resWaterPct")) $("resWaterPct").textContent = `${res.waterPct.toFixed(2)}%`;
    if ($("resSedimentPct")) $("resSedimentPct").textContent = `${res.sedimentPct.toFixed(2)}%`;
    if ($("resAcidPct")) $("resAcidPct").textContent = `${res.acidPct.toFixed(2)}%`;
    if ($("resDeductionPerTon")) $("resDeductionPerTon").textContent = String(res.deductionPerTon);
  }

  function clearLabInputs() {
    ["labClientName", "labAcids", "labEmptyWeight", "labBeforeBoil", "labAfterBoil", "labSedimentWeight"].forEach((id) => {
      if ($(id)) $(id).value = "";
    });
    hasCalculated = false;
    updateDisplay();
    window.showAppToast?.("تم تفريغ حقول التحليل.", "success");
  }

  function drawRoundedRect(ctx, x, y, w, h, r) {
    const radius = Math.min(r, w / 2, h / 2);
    ctx.beginPath();
    ctx.moveTo(x + radius, y);
    ctx.arcTo(x + w, y, x + w, y + h, radius);
    ctx.arcTo(x + w, y + h, x, y + h, radius);
    ctx.arcTo(x, y + h, x, y, radius);
    ctx.arcTo(x, y, x + w, y, radius);
    ctx.closePath();
  }

  function fmtWeightDisplay(v) {
    return Number.isInteger(v) ? String(v) : Number(v).toFixed(2).replace(/\.?0+$/, "");
  }

  async function shareAnalysisImage() {
    if (!hasCalculated) {
      window.showAppToast?.("اضغط احسب التحليل أولًا قبل المشاركة.", "warning");
      return;
    }
    const res = computeAnalysis();
    const sharedAt = new Date();
    const sharedDate = `${String(sharedAt.getDate()).padStart(2, "0")}-${String(sharedAt.getMonth() + 1).padStart(2, "0")}-${sharedAt.getFullYear()}`;
    const hour = sharedAt.getHours();
    const sharedTime = `${hour % 12 || 12}:${String(sharedAt.getMinutes()).padStart(2, "0")} ${hour >= 12 ? "م" : "ص"}`;

    // أبعاد عرضية (نسخة كمبيوتر)
    const canvas = document.createElement("canvas");
    canvas.width = 1600;
    canvas.height = 1045;
    const ctx = canvas.getContext("2d");

    ctx.direction = "rtl";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    // خلفية رصاصي ناعم
    ctx.fillStyle = "#e5e7eb";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // إطار رئيسي أبيض بحافة نبيتي
    ctx.fillStyle = "#ffffff";
    drawRoundedRect(ctx, 16, 16, canvas.width - 32, canvas.height - 32, 32);
    ctx.fill();
    ctx.lineWidth = 4;
    ctx.strokeStyle = "#881337";
    ctx.stroke();

    // العنوان الرئيسي (تحليل عينة) - لون رصاصي داكن
    const hasClient = (($("labClientName")?.value || "").trim().length > 0);
    ctx.fillStyle = "#1f2937";
    ctx.font = "800 56px 'Segoe UI', Tahoma, sans-serif";
    ctx.fillText("تحليل عينة", canvas.width / 2, 76);

    // اسم العميل - بلون مميز (نبيتي)
    if (hasClient) {
      ctx.fillStyle = "#881337";
      ctx.font = "900 50px 'Segoe UI', Tahoma, sans-serif";
      ctx.fillText(`(${res.clientName})`, canvas.width / 2, 138);
    }

    // سطر البيانات (المندوب | التاريخ | الوقت)
    const repName = String(window.getActiveUserName?.() || "").trim();
    ctx.fillStyle = "#6b7280";
    ctx.font = "600 30px 'Segoe UI', Tahoma, sans-serif";
    const shareDetails = `${repName ? `المندوب: ${repName}  |  ` : ""}التاريخ: ${sharedDate}  |  الوقت: ${sharedTime}`;
    ctx.fillText(shareDetails, canvas.width / 2, hasClient ? 196 : 165);

    // خط فاصل رصاصي
    const dividerY = hasClient ? 234 : 210;
    ctx.strokeStyle = "#d1d5db";
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(140, dividerY);
    ctx.lineTo(canvas.width - 140, dividerY);
    ctx.stroke();

    // ====== مجموعة الأوزان ======
    const wGroupX = 60;
    const wGroupY = dividerY + 26;
    const wGroupW = canvas.width - 120;
    const wGroupH = 235;

    ctx.fillStyle = "#f3f4f6";
    ctx.strokeStyle = "#d1d5db";
    ctx.lineWidth = 2;
    drawRoundedRect(ctx, wGroupX, wGroupY, wGroupW, wGroupH, 24);
    ctx.fill();
    ctx.stroke();

    const weights = [
      { label: "وزن فارغ", val: fmtWeightDisplay(res.emptyWeight) },
      { label: "قبل الغليان", val: fmtWeightDisplay(res.beforeBoil) },
      { label: "بعد الغليان", val: fmtWeightDisplay(res.afterBoil) },
      { label: "وزن الرواسب", val: fmtWeightDisplay(res.sedimentWeight) }
    ];
    const wGap = 20;
    const wPad = 24;
    const wCardW = (wGroupW - wPad * 2 - wGap * 3) / 4;
    const wCardH = wGroupH - wPad * 2;

    weights.forEach((item, idx) => {
      const cardX = wGroupX + wGroupW - wPad - (idx + 1) * wCardW - idx * wGap;
      const cardY = wGroupY + wPad;

      ctx.fillStyle = "#ffffff";
      ctx.strokeStyle = "#881337";
      ctx.lineWidth = 2;
      drawRoundedRect(ctx, cardX, cardY, wCardW, wCardH, 16);
      ctx.fill();
      ctx.stroke();

      ctx.fillStyle = "#4b5563";
      ctx.font = "700 26px 'Segoe UI', Tahoma, sans-serif";
      ctx.fillText(item.label, cardX + wCardW / 2, cardY + 42);

      ctx.fillStyle = "#1f2937";
      ctx.font = "900 48px 'Segoe UI', Tahoma, sans-serif";
      ctx.fillText(item.val, cardX + wCardW / 2, cardY + 115);
    });

    // ====== مجموعة النسب ======
    const rGroupX = 60;
    const rGroupY = wGroupY + wGroupH + 22;
    const rGroupW = canvas.width - 120;
    const rGroupH = 235;

    ctx.fillStyle = "#f3f4f6";
    ctx.strokeStyle = "#d1d5db";
    ctx.lineWidth = 2;
    drawRoundedRect(ctx, rGroupX, rGroupY, rGroupW, rGroupH, 24);
    ctx.fill();
    ctx.stroke();

    const ratios = [
      { label: "نسبة المياه", val: `${res.waterPct.toFixed(2)}%` },
      { label: "نسبة الرواسب", val: `${res.sedimentPct.toFixed(2)}%` },
      { label: "نسبة الأحماض", val: `${res.acidPct.toFixed(2)}%` }
    ];
    const rGap = 20;
    const rPad = 24;
    const rCardW = (rGroupW - rPad * 2 - rGap * 2) / 3;
    const rCardH = rGroupH - rPad * 2;

    ratios.forEach((item, idx) => {
      const cardX = rGroupX + rGroupW - rPad - (idx + 1) * rCardW - idx * rGap;
      const cardY = rGroupY + rPad;

      ctx.fillStyle = "#ffffff";
      ctx.strokeStyle = "#881337";
      ctx.lineWidth = 2;
      drawRoundedRect(ctx, cardX, cardY, rCardW, rCardH, 16);
      ctx.fill();
      ctx.stroke();

      ctx.fillStyle = "#4b5563";
      ctx.font = "700 28px 'Segoe UI', Tahoma, sans-serif";
      ctx.fillText(item.label, cardX + rCardW / 2, cardY + 45);

      ctx.fillStyle = "#881337";
      ctx.font = "900 54px 'Segoe UI', Tahoma, sans-serif";
      ctx.fillText(item.val, cardX + rCardW / 2, cardY + 125);
    });

    // ====== مربع الخصم ======
    const dedW = 540;
    const dedH = 148;
    const dedX = (canvas.width - dedW) / 2;
    const dedY = rGroupY + rGroupH + 26;

    ctx.fillStyle = "#881337";
    ctx.strokeStyle = "#4c0519";
    ctx.lineWidth = 3;
    drawRoundedRect(ctx, dedX, dedY, dedW, dedH, 24);
    ctx.fill();
    ctx.stroke();

    ctx.fillStyle = "#ffe4e6";
    ctx.font = "700 28px 'Segoe UI', Tahoma, sans-serif";
    ctx.fillText("الخصم (كجم/طن)", canvas.width / 2, dedY + 44);

    ctx.fillStyle = "#ffffff";
    ctx.font = "900 62px 'Segoe UI', Tahoma, sans-serif";
    ctx.fillText(String(res.deductionPerTon), canvas.width / 2, dedY + 106);

    // ====== التوقيع السفلي المنسق ======
    const footerY = dedY + dedH + 42;
    ctx.direction = "ltr";
    ctx.fillStyle = "#6b7280";
    ctx.font = "700 22px 'Segoe UI', Tahoma, sans-serif";
    ctx.fillText("Structured by ALI M. RAGAB  •  TAGADDOD", canvas.width / 2, footerY);

    // توليد الصورة والمشاركة
    const dataUrl = canvas.toDataURL("image/png");
    try {
      const blob = await (await fetch(dataUrl)).blob();
      const file = new File([blob], `تحليل_${res.clientName.replace(/\s+/g, "_")}.png`, { type: "image/png" });
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({
          files: [file],
          title: "تحليل عينة زيت"
        });
        window.showAppToast?.("تمت مشاركة صورة التحليل بنجاح.", "success");
        return;
      }
    } catch (_) {}

    const link = document.createElement("a");
    link.href = dataUrl;
    link.download = `تحليل_${res.clientName.replace(/\s+/g, "_")}.png`;
    link.click();
    window.showAppToast?.("تم حفظ صورة التحليل على الجهاز.", "success");
  }

  document.addEventListener("DOMContentLoaded", () => {
    $("calculateLabBtn")?.addEventListener("click", (event) => {
      hasCalculated = true;
      updateDisplay();
      withAnalysisButtonProgress(event.currentTarget, "تم حساب التحليل", () => {
        window.showAppToast?.("تم حساب نتائج التحليل.", "success");
        return true;
      });
    });
    $("shareLabImageBtn")?.addEventListener("click", (event) => {
      withAnalysisButtonProgress(event.currentTarget, "جاري تجهيز صورة التحليل...", shareAnalysisImage);
    });
    $("clearLabFormBtn")?.addEventListener("click", clearLabInputs);
    updateDisplay();
  });
})();