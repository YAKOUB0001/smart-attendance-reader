// app.js
// -----------------------------------------------------------------------------
// منطق الواجهة الأمامية. كل عملية فعلية (إنشاء الملف، فتح نوافذ اختيار
// الملفات، فتح الإكسل) تتم عبر استدعاء دوال بايثون المكشوفة في app_window.py
// من خلال الكائن الجاهز تلقائيًا: window.pywebview.api
// -----------------------------------------------------------------------------

const els = {
  monthSelect: document.getElementById("monthSelect"),
  yearInput: document.getElementById("yearInput"),
  nameInput: document.getElementById("nameInput"),
  phoneInput: document.getElementById("phoneInput"),
  folderPath: document.getElementById("folderPath"),
  browseBtn: document.getElementById("browseBtn"),
  openToggle: document.getElementById("openToggle"),
  createBtn: document.getElementById("createBtn"),
  dropzone: document.getElementById("dropzone"),
  previewWrap: document.getElementById("previewWrap"),
  previewImg: document.getElementById("previewImg"),
  imageName: document.getElementById("imageName"),
  replaceBtn: document.getElementById("replaceBtn"),
  stamp: document.getElementById("stamp"),
  stampMonth: document.getElementById("stampMonth"),
  stampYear: document.getElementById("stampYear"),
  statusBar: document.getElementById("statusBar"),
  confirmOverlay: document.getElementById("confirmOverlay"),
  confirmMessage: document.getElementById("confirmMessage"),
  confirmOk: document.getElementById("confirmOk"),
  confirmCancel: document.getElementById("confirmCancel"),
  // حقول الراتب الجديدة
  baseRateInput: document.getElementById("baseRateInput"),
  allowanceRateInput: document.getElementById("allowanceRateInput"),
  advancesInput: document.getElementById("advancesInput"),
  totalRateDisplay: document.getElementById("totalRateDisplay"),
};

let statusTimer = null;

// -----------------------------------------------------------------------------
// عرض رسالة حالة عائمة في أسفل النافذة (نجاح / خطأ / معلومة عادية)
// -----------------------------------------------------------------------------
function showStatus(message, kind = "info") {
  els.statusBar.textContent = message;
  els.statusBar.className = "status show " + (kind === "info" ? "" : kind);
  clearTimeout(statusTimer);
  statusTimer = setTimeout(() => {
    els.statusBar.classList.remove("show");
  }, 4500);
}

// -----------------------------------------------------------------------------
// تحديث ختم الشهر/السنة أعلى النافذة مع حركة بسيطة عند التغيير
// -----------------------------------------------------------------------------
function updateStamp() {
  const monthText = els.monthSelect.options[els.monthSelect.selectedIndex]?.text || "—";
  els.stampMonth.textContent = monthText;
  els.stampYear.textContent = els.yearInput.value || "—";
  els.stamp.classList.remove("stamp-pop");
  void els.stamp.offsetWidth;
  els.stamp.classList.add("stamp-pop");
}

// -----------------------------------------------------------------------------
// تحديث عرض "إجمالي سعر الساعة" مباشرة عند تغيير سعر الساعة أو العلاوة
// -----------------------------------------------------------------------------
function updateTotalRateDisplay() {
  const base = parseFloat(els.baseRateInput.value) || 0;
  const allowance = parseFloat(els.allowanceRateInput.value) || 0;
  els.totalRateDisplay.textContent = (base + allowance).toFixed(2);
}

// -----------------------------------------------------------------------------
// نافذة تأكيد مخصصة (بديل نافذة النظام) - ترجع Promise<boolean>
// -----------------------------------------------------------------------------
function askConfirm(message) {
  return new Promise((resolve) => {
    els.confirmMessage.textContent = message;
    els.confirmOverlay.hidden = false;

    const cleanup = (result) => {
      els.confirmOverlay.hidden = true;
      els.confirmOk.removeEventListener("click", onOk);
      els.confirmCancel.removeEventListener("click", onCancel);
      resolve(result);
    };
    const onOk = () => cleanup(true);
    const onCancel = () => cleanup(false);

    els.confirmOk.addEventListener("click", onOk);
    els.confirmCancel.addEventListener("click", onCancel);
  });
}

// -----------------------------------------------------------------------------
// تهيئة الصفحة بعد أن يصبح جسر بايثون (pywebview.api) جاهزًا
// -----------------------------------------------------------------------------
async function initApp() {
  const months = await window.pywebview.api.get_months();
  months.forEach((name, index) => {
    const opt = document.createElement("option");
    opt.value = index + 1;
    opt.textContent = name;
    els.monthSelect.appendChild(opt);
  });

  const defaults = await window.pywebview.api.get_defaults();
  els.yearInput.value = defaults.year;
  els.monthSelect.value = defaults.month;
  els.folderPath.textContent = defaults.outputDir;
  els.folderPath.title = defaults.outputDir;
  els.folderPath.dataset.fullPath = defaults.outputDir;

  // القيم الافتراضية لحساب الراتب (قادمة من الخلفية حتى تبقى مصدرًا واحدًا للحقيقة)
  els.baseRateInput.value = defaults.baseRate;
  els.allowanceRateInput.value = defaults.allowanceRate;
  els.advancesInput.value = defaults.advances;

  updateStamp();
  updateTotalRateDisplay();
}

// -----------------------------------------------------------------------------
// اختيار صورة جدول الحضور ومعاينتها
// -----------------------------------------------------------------------------
async function pickImage() {
  const result = await window.pywebview.api.choose_image();
  if (!result) return;
  if (result.error) {
    showStatus("تعذّر فتح الصورة: " + result.error, "error");
    return;
  }
  els.previewImg.src = result.dataUrl;
  els.imageName.textContent = result.name;
  els.imageName.title = result.name;
  els.dropzone.hidden = true;
  els.previewWrap.hidden = false;
  showStatus("تم اختيار الصورة. (معالجة OCR ستُضاف في الخطوة القادمة)", "success");
}

// -----------------------------------------------------------------------------
// إنشاء كشف الحضور والراتب
// -----------------------------------------------------------------------------
async function createAttendance() {
  const options = {
    month: parseInt(els.monthSelect.value, 10),
    year: parseInt(els.yearInput.value, 10),
    name: els.nameInput.value.trim(),
    phone: els.phoneInput.value.trim(),
    folder: els.folderPath.dataset.fullPath || els.folderPath.textContent,
    openAfter: els.openToggle.checked,
    // بيانات الراتب الجديدة تُرسل كـ JSON إلى الخلفية
    baseRate: parseFloat(els.baseRateInput.value) || 0,
    allowanceRate: parseFloat(els.allowanceRateInput.value) || 0,
    advances: parseFloat(els.advancesInput.value) || 0,
  };

  if (!options.month || !options.year) {
    showStatus("الرجاء اختيار الشهر والسنة أولًا.", "error");
    return;
  }

  if (options.baseRate <= 0) {
    showStatus("الرجاء إدخال سعر ساعة أساسي أكبر من صفر.", "error");
    return;
  }

  els.createBtn.disabled = true;
  try {
    const existing = await window.pywebview.api.check_exists(options.folder, options.year, options.month);
    if (existing && existing.exists) {
      const confirmed = await askConfirm(
        `يوجد ملف بنفس الاسم:\n${existing.fileName}\nهل تريد استبداله؟`
      );
      if (!confirmed) {
        showStatus("تم الإلغاء. لم يتم إنشاء أي ملف.");
        return;
      }
    }

    const result = await window.pywebview.api.create_attendance(options);
    if (result.success) {
      showStatus("✅ تم إنشاء الملف بنجاح: " + result.fileName, "success");
    } else {
      showStatus("⚠️ " + result.error, "error");
    }
  } catch (err) {
    showStatus("⚠️ حدث خطأ غير متوقع: " + err, "error");
  } finally {
    els.createBtn.disabled = false;
  }
}

// -----------------------------------------------------------------------------
// اختيار مجلد الحفظ
// -----------------------------------------------------------------------------
async function pickFolder() {
  const folder = await window.pywebview.api.choose_folder();
  if (folder) {
    els.folderPath.textContent = folder;
    els.folderPath.title = folder;
    els.folderPath.dataset.fullPath = folder;
  }
}

// -----------------------------------------------------------------------------
// ربط الأحداث
// -----------------------------------------------------------------------------
els.dropzone.addEventListener("click", pickImage);
els.dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") pickImage();
});
els.replaceBtn.addEventListener("click", pickImage);
els.browseBtn.addEventListener("click", pickFolder);
els.createBtn.addEventListener("click", createAttendance);
els.monthSelect.addEventListener("change", updateStamp);
els.yearInput.addEventListener("input", updateStamp);
els.baseRateInput.addEventListener("input", updateTotalRateDisplay);
els.allowanceRateInput.addEventListener("input", updateTotalRateDisplay);

["dragenter", "dragover"].forEach((evt) =>
  els.dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    els.dropzone.classList.add("drag-over");
  })
);
["dragleave", "drop"].forEach((evt) =>
  els.dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    els.dropzone.classList.remove("drag-over");
    if (evt === "drop") pickImage();
  })
);

// -----------------------------------------------------------------------------
// نقطة البداية: ننتظر جسر بايثون قبل أي استدعاء
// -----------------------------------------------------------------------------
if (window.pywebview) {
  initApp();
} else {
  window.addEventListener("pywebviewready", initApp);
}