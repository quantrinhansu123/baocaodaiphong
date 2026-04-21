/**
 * Báo cáo tồn kho nhiên liệu — AppSheet.
 * «Mã nhiên liệu» / «Tên nhiên liệu» lấy từ «Danh sách tài sản»
 * (cột «Mã tài sản», ưu tiên «Tên nhiên liệu», không có thì «Tên tài sản») khi «Tên nhóm» = «Nhiên liệu».
 * «Tồn đầu kỳ» (SL, và GT nếu có): bảng «Tồn kho ĐK» — cột «Sl tồn ĐK»,
 * khớp «Tên NL» với tên nhiên liệu và «Ngày nhập số dư ĐK» nằm trong khoảng «Từ ngày»..«Tới ngày».
 * «Nhập trong kỳ» / «Xuất trong kỳ» (SL): «Nhập xuất luân chuyển CT» — tổng «Số lượng NL»;
 * Giá trị nhập/xuất kỳ: lấy từ cột «Thành tiền NL» trên «Nhập xuất luân chuyển CT».
 * Nhập / Xuất trong kỳ: chỉ theo cột «Loại phiếu» = Nhập hoặc Xuất (cùng bộ lọc ngày & NCC / nơi nhập / xe).
 * Các ô SL (tồn/nhập/xuất/tồn cuối) hiển thị kèm «(ĐVT)» từ cột ĐVT trên «Danh sách tài sản» của dòng đó.
 */
(function () {
    "use strict";

    const APPSHEET_CONFIG = {
        appId: "3be5baea-960f-4d3f-b388-d13364cc4f22",
        accessKey: "V2-GaoRd-ItaM1-r44oH-c6Smd-uOe7V-cmVoK-IJINF-5XLQa"
    };

    const TABLE_DS_TAI_SAN = "Danh sách tài sản";
    const TABLE_TON_KHO_DK = "Tồn kho ĐK";
    const TABLE_NXLC_CT = "Nhập xuất luân chuyển CT";
    const NHOM_NHIEN_LIEU_LABEL = "Nhiên liệu";

    let cacheDsRows = null;
    let cacheTonKhoDkRows = null;
    /** Toàn bộ «Nhập xuất luân chuyển CT» (null = chưa tải). */
    let cacheNxlcCtRows = null;
    /** Có lỗi HTTP/AppSheet khi tải NXLC — nhập kỳ quay về cột trên DS. */
    let lastNxlcFetchError = null;
    /** Dòng NL (sau lọc DS) — dùng render lại khi đổi «Từ ngày». */
    let cacheFuelRows = [];

    function cellString(v) {
        if (v == null) return "";
        if (typeof v === "object" && v !== null && "Value" in v) return String(v.Value);
        if (typeof v === "object" && v !== null && "value" in v) return String(v.value);
        return String(v);
    }

    function cellDisplayString(v) {
        if (v == null) return "";
        if (typeof v === "object" && v !== null) {
            const display = v.DisplayValue ?? v.displayValue ?? v.FormattedValue ?? v.formattedValue ?? v.Text ?? v.text;
            if (display != null && String(display).trim() !== "") return String(display);
        }
        return cellString(v);
    }

    function normalizeAppSheetDateCell(value, keyHint) {
        const keyLooksDate = /ng[aà]y|date/i.test(String(keyHint ?? ""));
        const normalizeString = (s) => {
            const text = String(s ?? "").trim();
            if (!text) return s;
            if (!keyLooksDate && !/^\d{4}[/-]\d/.test(text) && !/^\d{1,2}[/-]\d/.test(text)) return s;
            return s;
        };
        if (typeof value === "string") return normalizeString(value);
        if (!value || typeof value !== "object") return value;
        const out = { ...value };
        for (const field of ["DisplayValue", "displayValue", "FormattedValue", "formattedValue", "Text", "text", "Value", "value"]) {
            if (out[field] == null || typeof out[field] !== "string") continue;
            out[field] = normalizeString(out[field]);
        }
        return out;
    }

    function normalizeAppSheetRowDates(row) {
        if (!row || typeof row !== "object") return row;
        const out = { ...row };
        for (const [k, v] of Object.entries(out)) {
            out[k] = normalizeAppSheetDateCell(v, k);
        }
        return out;
    }

    async function fetchAppSheetTable(tableName) {
        const url = `https://api.appsheet.com/api/v2/apps/${encodeURIComponent(APPSHEET_CONFIG.appId)}/tables/${encodeURIComponent(tableName)}/Action`;
        const payload = {
            Action: "Find",
            Properties: { Locale: "vi-VN", Timezone: "Asia/Ho_Chi_Minh" },
            Rows: []
        };
        const response = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                ApplicationAccessKey: APPSHEET_CONFIG.accessKey
            },
            body: JSON.stringify(payload)
        });
        if (!response.ok) {
            const err = await response.text();
            throw new Error(`${tableName}: ${response.status} ${err}`);
        }
        const data = await response.json();
        if (Array.isArray(data)) return data.map(normalizeAppSheetRowDates);
        if (data && Array.isArray(data.Rows)) return data.Rows.map(normalizeAppSheetRowDates);
        return [];
    }

    function parseNumeric(raw) {
        if (raw == null || raw === "") return null;
        if (typeof raw === "number" && !isNaN(raw)) return raw;
        const s0 = cellDisplayString(raw).trim();
        const m = s0.match(/-?[\d.,]+/);
        if (!m) return null;
        let s = m[0];
        const comma = s.indexOf(",");
        const dot = s.indexOf(".");
        if (comma !== -1 && dot !== -1) {
            if (comma > dot) s = s.replace(/\./g, "").replace(",", ".");
            else s = s.replace(/,/g, "");
        } else if (comma !== -1) {
            s = s.replace(",", ".");
        } else if (dot !== -1 && s.indexOf(".", dot + 1) !== -1) {
            s = s.replace(/\./g, "");
        }
        const n = parseFloat(s);
        return isNaN(n) ? null : n;
    }

    function formatNum(n) {
        if (n == null || isNaN(n)) return "";
        return new Intl.NumberFormat("vi-VN", { minimumFractionDigits: 0, maximumFractionDigits: 2 }).format(n);
    }

    function escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    function pickFirstCell(row, keys) {
        for (const k of keys) {
            const v = row[k];
            if (v === undefined || v === null) continue;
            const s = cellDisplayString(v).trim();
            if (s !== "") return s;
        }
        return "";
    }

    function pickDsMaThietBi(row) {
        return pickFirstCell(row, [
            "Mã tài sản",
            "Ma tai san",
            "Mã thiết bị",
            "Ma thiet bi",
            "Mã TB",
            "Ma TB",
            "Mã",
            "Ma"
        ]);
    }

    function pickDsTenThietBi(row) {
        return pickFirstCell(row, [
            "Tên tài sản",
            "Ten tai san",
            "Tên thiết bị",
            "Ten thiet bi",
            "TenTaiSan",
            "TenThietBi"
        ]);
    }

    /** Ưu tiên cột «Tên nhiên liệu» trên AppSheet; fallback «Tên tài sản». */
    function pickDsTenNhienLieu(row) {
        const direct = pickFirstCell(row, [
            "Tên nhiên liệu",
            "Ten nhien lieu",
            "TenNhienLieu",
            "Tên NL",
            "Ten NL"
        ]);
        if (direct) return direct;
        return pickDsTenThietBi(row);
    }

    function pickDsDvt(row) {
        return pickFirstCell(row, ["ĐVT", "DVT", "Đơn vị tính", "Don vi tinh", "Đơn vị", "Don vi"]);
    }

    function pickDsNhomThietBi(row) {
        return pickFirstCell(row, [
            "Nhóm thiết bị",
            "Nhom thiet bi",
            "Tên nhóm",
            "Ten nhom",
            "Nhóm xe máy thiết bị",
            "Nhom xe may thiet bi"
        ]);
    }

    /** Giá trị «Tên nhóm» trên DS — ưu tiên đúng tên cột trước «Nhóm thiết bị». */
    function pickDsTenNhom(row) {
        const tn = pickFirstCell(row, ["Tên nhóm", "Ten nhom", "TenNhom"]);
        if (tn) return tn;
        return pickDsNhomThietBi(row);
    }

    function pickDsGhiChu(row) {
        return pickFirstCell(row, ["Ghi chú", "Ghi chu", "GhiChu", "Mô tả", "Mo ta"]);
    }

    function pickDsKhoanMucChiPhi(row) {
        return pickFirstCell(row, [
            "Khoản mục chi phí",
            "Khoan muc chi phi",
            "Khoản mục CP",
            "Loại chi phí",
            "Loai chi phi",
            "Mục chi phí",
            "Muc chi phi"
        ]);
    }

    function normalizeKeyPart(s) {
        return String(s ?? "")
            .trim()
            .normalize("NFD")
            .replace(/\u0300-\u036f/g, "")
            .toLowerCase()
            .replace(/[^a-z0-9]/g, "");
    }

    function rowIsNhomNhienLieu(row) {
        const g = pickDsTenNhom(row).trim();
        if (!g) return false;
        return normalizeKeyPart(g) === normalizeKeyPart(NHOM_NHIEN_LIEU_LABEL);
    }

    function buildStrictDate(y, mm, dd) {
        if (!y || mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;
        const d = new Date(y, mm - 1, dd);
        if (isNaN(d.getTime())) return null;
        if (d.getFullYear() !== y || d.getMonth() !== mm - 1 || d.getDate() !== dd) return null;
        return d;
    }

    function parseDateInputYmd(value) {
        if (!value) return null;
        const m = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (!m) return null;
        const y = parseInt(m[1], 10);
        const mm = parseInt(m[2], 10);
        const dd = parseInt(m[3], 10);
        return buildStrictDate(y, mm, dd);
    }

    /** Giống BaocaoNL-appsheet — ô «dd/mm/yyyy» + lịch native. */
    function parseDateFlexible(value) {
        if (value instanceof Date) return value;
        const ymd = parseDateInputYmd(value);
        if (ymd) return ymd;
        const s = String(value ?? "").trim();
        if (!s) return null;
        let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
        if (m) return buildStrictDate(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));
        m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[\s,].*)?$/);
        if (m) {
            const a = parseInt(m[1], 10);
            const b = parseInt(m[2], 10);
            const y = parseInt(m[3], 10);
            if (a <= 12 && b <= 12 && /\b(am|pm)\b/i.test(s)) return buildStrictDate(y, a, b);
            return buildStrictDate(y, b, a);
        }
        m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})(?:[\s,].*)?$/);
        if (m) {
            const a = parseInt(m[1], 10);
            const b = parseInt(m[2], 10);
            const y = parseInt(m[3], 10);
            return buildStrictDate(y, b, a);
        }
        if (/[a-z]/i.test(s)) {
            const native = new Date(s);
            if (!isNaN(native.getTime())) {
                return new Date(native.getFullYear(), native.getMonth(), native.getDate());
            }
        }
        return null;
    }

    function formatDateVn(value) {
        const d = parseDateFlexible(value);
        if (!d || isNaN(d.getTime())) return "";
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const yyyy = d.getFullYear();
        return `${dd}/${mm}/${yyyy}`;
    }

    function dateToYmdInputValue(d) {
        if (!d || isNaN(d.getTime())) return "";
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, "0");
        const day = String(d.getDate()).padStart(2, "0");
        return `${y}-${m}-${day}`;
    }

    function wireNativeDatePickerButton(textId, pickerId, buttonId) {
        const textEl = document.getElementById(textId);
        const pickerEl = document.getElementById(pickerId);
        const btnEl = document.getElementById(buttonId);
        if (!textEl || !pickerEl || !btnEl) return;
        btnEl.addEventListener("click", (e) => {
            e.preventDefault();
            const d = parseDateFlexible(textEl.value);
            pickerEl.value = d ? dateToYmdInputValue(d) : "";
            try {
                if (typeof pickerEl.showPicker === "function") {
                    pickerEl.showPicker();
                } else {
                    pickerEl.focus();
                    pickerEl.click();
                }
            } catch (_) {
                pickerEl.click();
            }
        });
        pickerEl.addEventListener("change", () => {
            const v = pickerEl.value;
            if (v) {
                const parsed = parseDateInputYmd(v);
                textEl.value = parsed ? formatDateVn(parsed) : "";
            } else {
                textEl.value = "";
            }
            textEl.dispatchEvent(new Event("input", { bubbles: true }));
            textEl.dispatchEvent(new Event("change", { bubbles: true }));
        });
    }

    /** Khi chưa chọn công trình cụ thể — giống BaocaoNL (mặc định hai công trường). */
    const TKH_DEFAULT_BANNER_SITE_LINE = "CT HỐ SÓI - SÔNG ĐUỐNG";

    function syncTkhDateRangeBanner() {
        const dr = document.getElementById("tkh-date-range");
        if (!dr) return;
        const fromRaw = document.getElementById("tkh-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("tkh-filter-to-date")?.value?.trim() ?? "";
        const dFrom = fromRaw ? parseDateFlexible(fromRaw) : null;
        const dTo = toRaw ? parseDateFlexible(toRaw) : null;
        const labelFrom =
            !fromRaw ? "…" : dFrom && !isNaN(dFrom.getTime()) ? formatDateVn(dFrom) : fromRaw;
        const labelTo = !toRaw ? "…" : dTo && !isNaN(dTo.getTime()) ? formatDateVn(dTo) : toRaw;
        dr.textContent = `(Từ ngày ${labelFrom} tới ngày ${labelTo})`;
    }

    /** Dòng «CT …» dưới tiêu đề: theo giá trị «Công trình»; «Tất cả» = liệt kê các công trình có trong dữ liệu hoặc mặc định. */
    function syncTkhCongTrinhBanner() {
        const el = document.getElementById("tkh-banner-cong-trinh-line");
        const sel = document.getElementById("tkh-filter-cong-trinh");
        if (!el || !sel) return;
        const v = sel.value?.trim() ?? "";
        if (v) {
            el.textContent = v;
            return;
        }
        const options = [...sel.options]
            .slice(1)
            .map((o) => String(o.value ?? "").trim())
            .filter(Boolean);
        if (!options.length) {
            el.textContent = TKH_DEFAULT_BANNER_SITE_LINE;
            return;
        }
        if (options.length === 1) {
            el.textContent = options[0];
            return;
        }
        el.textContent = options.join(" • ");
    }

    function syncTkhBanners() {
        syncTkhDateRangeBanner();
        syncTkhCongTrinhBanner();
    }

    function parseCellToDate(raw) {
        if (raw == null || raw === "") return null;
        if (raw instanceof Date && !isNaN(raw.getTime())) return raw;
        const s = cellDisplayString(raw).trim();
        if (!s) return null;
        const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:T|[\sZ]|$)/);
        if (iso) {
            const y = parseInt(iso[1], 10);
            const m = parseInt(iso[2], 10);
            const day = parseInt(iso[3], 10);
            return buildStrictDate(y, m, day);
        }
        const slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (slash) {
            const a = parseInt(slash[1], 10);
            const b = parseInt(slash[2], 10);
            const y = parseInt(slash[3], 10);
            if (a > 12 && b <= 12) return buildStrictDate(y, b, a); // d/m/yyyy (không mơ hồ)
            if (a <= 12 && b > 12) return buildStrictDate(y, a, b); // m/d/yyyy (không mơ hồ)
            // Trường hợp mơ hồ (vd 4/9/2026): dữ liệu AppSheet thường là m/d/yyyy.
            return buildStrictDate(y, a, b);
        }
        const d = new Date(s);
        if (!isNaN(d.getTime())) return d;
        return null;
    }

    function dateToDdMmYyyy(d) {
        if (!d || isNaN(d.getTime())) return "";
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        return `${dd}/${mm}/${d.getFullYear()}`;
    }

    function anyValueToDdMmYyyy(v) {
        const d = parseCellToDate(v);
        return d ? dateToDdMmYyyy(d) : "";
    }

    /** Ép hiển thị ngày trong modal xem trưới — thử parseCellToDate rồi parseDateFlexible trên chuỗi hiển thị. */
    function cellValueToDdMmYyyyStr(v) {
        if (v == null || v === "") return "";
        let d = parseCellToDate(v);
        if ((!d || isNaN(d.getTime())) && !(v instanceof Date)) {
            const s = cellDisplayString(v).trim();
            if (s) {
                const df = parseDateFlexible(s);
                if (df && !isNaN(df.getTime())) d = df;
            }
        }
        if (!d || isNaN(d.getTime())) return "";
        return dateToDdMmYyyy(d);
    }

    function isLikelyDateTimeColumnKey(k) {
        const s = String(k);
        const n = s
            .normalize("NFD")
            .replace(/\u0300-\u036f/g, "")
            .toLowerCase();
        if (/^ngay\b|^date\b|^time\b|ngay_|_ngay$/i.test(s)) return true;
        if (/\bngay\b/.test(n)) return true;
        if (/\bdate\b|\btime\b|thoi\s*gian|gio\s*tai|created|modified/i.test(n)) return true;
        return false;
    }

    function looksLikeRawDateString(text) {
        const t = String(text ?? "").trim();
        if (!t) return false;
        if (/^\d{4}-\d{2}-\d{2}/.test(t)) return true;
        if (/^\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}/.test(t)) return true;
        if (/\d{4}-\d{2}-\d{2}T/.test(t)) return true;
        return false;
    }

    function getFromDateDdMmYyyyFromDom() {
        const el = document.getElementById("tkh-filter-from-date");
        const v = el?.value?.trim() ?? "";
        if (!v) return "";
        const d = parseDateFlexible(v);
        return d ? dateToDdMmYyyy(d) : "";
    }

    function getSelectedDateRangeFromDom() {
        const fromRaw = document.getElementById("tkh-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("tkh-filter-to-date")?.value?.trim() ?? "";
        let fromD = fromRaw ? parseDateFlexible(fromRaw) : null;
        let toD = toRaw ? parseDateFlexible(toRaw) : null;
        if (!fromD && !toD) return null;
        if (!fromD) fromD = toD;
        if (!toD) toD = fromD;
        let start = new Date(fromD.getFullYear(), fromD.getMonth(), fromD.getDate(), 0, 0, 0, 0);
        let end = new Date(toD.getFullYear(), toD.getMonth(), toD.getDate(), 23, 59, 59, 999);
        if (start > end) {
            const tmp = start;
            start = end;
            end = tmp;
        }
        return { start, end };
    }

    /** «Tên NL» trên «Tồn kho ĐK» — khớp với «Tên nhiên liệu» (ưu tiên) trên «Danh sách tài sản». */
    function pickTkTenNl(row) {
        const direct = pickFirstCell(row, [
            "Tên NL",
            "Ten NL",
            "Tên nhiên liệu",
            "Ten nhien lieu",
            "Tên NL ",
            "Tên nhiên liệu "
        ]);
        if (direct) return direct;
        for (const [k, v] of Object.entries(row || {})) {
            if (/^t[eê]n\s*nl$/i.test(String(k).trim()) || /^ten\s*nl$/i.test(String(k).trim())) {
                const s = cellDisplayString(v).trim();
                if (s) return s;
            }
        }
        return "";
    }

    /** «Ngày nhập số dư ĐK» — so sánh sau khi ép dd/mm/yyyy. */
    function pickTkNgayNhapSoDu(row) {
        const keys = [
            "Ngày nhập số dư ĐK",
            "Ngay nhap so du DK",
            "Ngày nhập so du DK",
            "Ngày nhập số dư DK",
            "Ngay nhap so du ĐK",
            "Ngày nhập SD ĐK",
            "Ngay nhap SD DK"
        ];
        for (const k of keys) {
            if (row[k] === undefined || row[k] === null) continue;
            const s = cellDisplayString(row[k]).trim();
            if (s !== "") return row[k];
        }
        for (const [k, v] of Object.entries(row || {})) {
            const kn = String(k);
            if (/nh[aậ]p/i.test(kn) && /s[oố]\s*d[uư]/i.test(kn) && (/[Đd][Kk]|DK/i.test(kn) || /d[uư]\s*[ĐdKk]/i.test(kn))) {
                const s = cellDisplayString(v).trim();
                if (s !== "") return v;
            }
        }
        return null;
    }

    function pickNumericByAliases(row, exactKeys, keyRegexFallback) {
        for (const k of exactKeys) {
            if (row[k] == null || row[k] === "") continue;
            const n = parseNumeric(row[k]);
            if (n != null) return n;
        }
        if (keyRegexFallback) {
            for (const [k, v] of Object.entries(row || {})) {
                if (!keyRegexFallback.test(String(k))) continue;
                const n = parseNumeric(v);
                if (n != null) return n;
            }
        }
        return null;
    }

    function extractStockCells(row) {
        const tonDauSl = pickNumericByAliases(
            row,
            [
                "SL tồn đầu kỳ",
                "SL ton dau ky",
                "Tồn đầu kỳ (SL)",
                "Ton dau ky (SL)",
                "Tồn đầu SL",
                "SL tồn đầu",
                "Ton dau SL"
            ],
            /ton\s*dau|dau\s*ky.*sl|sl.*dau|ton\s*kho\s*dau/i
        );

        const tonDauGt = pickNumericByAliases(
            row,
            [
                "Giá trị tồn đầu kỳ",
                "Gia tri ton dau ky",
                "GT tồn đầu",
                "GT ton dau",
                "Giá trị tồn đầu",
                "Gia tri ton dau"
            ],
            /gia\s*tri.*ton\s*dau|gt.*ton\s*dau|ton\s*dau.*gia|gia.*dau\s*ky/i
        );

        const nhapSl = pickNumericByAliases(
            row,
            ["SL nhập trong kỳ", "SL nhap trong ky", "Nhập trong kỳ (SL)", "SL nhập kỳ", "Nhập SL"],
            /nhap.*ky.*sl|sl.*nhap|nhap\s*trong.*sl/i
        );

        const nhapGt = pickNumericByAliases(
            row,
            ["Giá trị nhập trong kỳ", "GT nhập trong kỳ", "GT nhập kỳ", "Giá trị nhập"],
            /gia\s*tri.*nhap|gt.*nhap|nhap.*gia/i
        );

        const xuatSl = pickNumericByAliases(
            row,
            ["SL xuất trong kỳ", "SL xuat trong ky", "Xuất trong kỳ (SL)", "SL xuất kỳ", "Xuất SL"],
            /xuat.*ky.*sl|sl.*xuat|xuat\s*trong.*sl/i
        );

        const xuatGt = pickNumericByAliases(
            row,
            ["Giá trị xuất trong kỳ", "GT xuất trong kỳ", "GT xuất kỳ", "Giá trị xuất"],
            /gia\s*tri.*xuat|gt.*xuat|xuat.*gia/i
        );

        const tonCuoiSl = pickNumericByAliases(
            row,
            [
                "SL tồn cuối kỳ",
                "SL ton cuoi ky",
                "Tồn cuối kỳ (SL)",
                "Tồn cuối SL",
                "SL tồn cuối"
            ],
            /ton\s*cuoi|cuoi\s*ky.*sl|sl.*cuoi/i
        );

        const tonCuoiGt = pickNumericByAliases(
            row,
            ["Giá trị tồn cuối kỳ", "GT tồn cuối", "Giá trị tồn cuối", "GT ton cuoi"],
            /gia\s*tri.*ton\s*cuoi|gt.*ton\s*cuoi|ton\s*cuoi.*gia/i
        );

        return {
            tonDauSl,
            tonDauGt,
            nhapSl,
            nhapGt,
            xuatSl,
            xuatGt,
            tonCuoiSl,
            tonCuoiGt
        };
    }

    function pickTkSlTonDk(row) {
        return pickNumericByAliases(
            row,
            [
                "Sl tồn ĐK",
                "SL tồn ĐK",
                "Sl tồn DK",
                "SL tồn DK",
                "SL tồn Đk",
                "Sl ton ĐK",
                "SL ton DK"
            ],
            /sl\s*t[oô]n\s*[Đd][Kk]|sl\s*ton\s*dk/i
        );
    }

    function pickTkGtTonDk(row) {
        return pickNumericByAliases(
            row,
            ["Giá trị tồn ĐK", "Gia tri ton ĐK", "GT tồn ĐK", "GT tồn DK", "Giá trị tồn DK"],
            /gia\s*tri.*t[oô]n\s*[Đd][Kk]|gt.*t[oô]n.*[Đd][Kk]/i
        );
    }

    /**
     * Gom «Tồn kho ĐK» theo khóa tên NL (chuẩn hóa).
     * - Luôn đọc từ cột «Tên NL» + «Sl tồn ĐK» (và GT tồn ĐK nếu có).
     * - Nếu có khoảng ngày, chỉ lấy các dòng nằm trong khoảng đó.
     * Trả về Map(normalizeKeyPart(Tên NL)) -> { sl, gt } (cộng dồn nếu trùng tên).
     */
    function buildTonKhoDkAggregateByTenNl(tonKhoRows, dateRange) {
        const map = new Map();
        if (!tonKhoRows?.length) return map;
        for (const r of tonKhoRows) {
            if (dateRange) {
                const rawNgay = pickTkNgayNhapSoDu(r);
                const rowDate = rawNgay != null ? parseCellToDate(rawNgay) : null;
                if (!rowDate || isNaN(rowDate.getTime())) continue;
                if (rowDate < dateRange.start || rowDate > dateRange.end) continue;
            }
            const tenNl = pickTkTenNl(r).trim();
            if (!tenNl) continue;
            const nk = normalizeKeyPart(tenNl);
            const sl = pickTkSlTonDk(r);
            const gt = pickTkGtTonDk(r);
            const prev = map.get(nk) || { slSum: 0, gtSum: 0, hasSl: false, hasGt: false };
            if (sl != null && !isNaN(sl)) {
                prev.slSum += sl;
                prev.hasSl = true;
            }
            if (gt != null && !isNaN(gt)) {
                prev.gtSum += gt;
                prev.hasGt = true;
            }
            map.set(nk, prev);
        }
        const out = new Map();
        for (const [k, v] of map) {
            out.set(k, {
                sl: v.hasSl ? v.slSum : null,
                gt: v.hasGt ? v.gtSum : null
            });
        }
        return out;
    }

    function loaiPhieuIsNhap(rawVal) {
        const s = String(rawVal ?? "")
            .trim()
            .toLowerCase()
            .normalize("NFD")
            .replace(/\u0300-\u036f/g, "");
        if (!s) return false;
        if (s === "nhap" || s.includes("phieu nhap") || s.includes("nhap kho") || s.includes("nhap hang")) return true;
        return false;
    }

    /**
     * Chỉ đọc cột «Loại phiếu» trên NXLC (Nhập / Xuất phải khớp đúng cột này).
     * Không dùng «Trạng thái», «Loại giao dịch», «Loại» đơn lẻ.
     */
    function getLoaiPhieuColumnNxlc(row) {
        if (!row || typeof row !== "object") return "";
        const exact = ["Loại phiếu", "Loai phieu", "Loại phiếu NX", "Loai phieu NX", "Loại phiếu NXLC"];
        for (const k of exact) {
            if (row[k] == null || String(row[k]).trim() === "") continue;
            return row[k];
        }
        for (const [k, v] of Object.entries(row)) {
            if (v == null || String(v).trim() === "") continue;
            const kn = String(k).trim();
            if (/^lo[aạ]i\s*phi[ếe]u(\s+nx|\s+nxlc)?$/i.test(kn) || /^loai\s*phieu(\s+nx)?$/i.test(kn)) return v;
        }
        return "";
    }

    /** Phiếu Nhập — chỉ theo giá trị cột «Loại phiếu». */
    function nxlcRowIsPhieuNhap(r) {
        const lp = getLoaiPhieuColumnNxlc(r);
        const sLoai = cellDisplayString(lp).trim();
        if (!sLoai) return false;
        if (sLoai === "Xuất" || sLoai === "XUẤT" || loaiPhieuIsXuat(sLoai)) return false;
        if (sLoai === "Nhập" || sLoai === "NHẬP") return true;
        return loaiPhieuIsNhap(sLoai);
    }

    function loaiPhieuIsXuat(rawVal) {
        const s = String(rawVal ?? "")
            .trim()
            .toLowerCase()
            .normalize("NFD")
            .replace(/\u0300-\u036f/g, "");
        if (!s) return false;
        if (s === "xuat") return true;
        if (s.startsWith("xuat ") || s.startsWith("xuat-") || s.startsWith("xuat/")) return true;
        if (s.includes("phieu xuat") || s.includes("xuat kho") || s.includes("xuat hang")) return true;
        return false;
    }

    /** Phiếu Xuất — chỉ theo giá trị cột «Loại phiếu». */
    function nxlcRowIsPhieuXuat(r) {
        const lp = getLoaiPhieuColumnNxlc(r);
        const sLoai = cellDisplayString(lp).trim();
        if (!sLoai) return false;
        if (sLoai === "Nhập" || sLoai === "NHẬP" || loaiPhieuIsNhap(sLoai)) return false;
        if (sLoai === "Xuất" || sLoai === "XUẤT") return true;
        return loaiPhieuIsXuat(sLoai);
    }

    /** Ngày chứng từ — ưu «Ngày»; phiếu nhập có thể chỉ có «Ngày nhập». */
    function pickNxlcNgay(row) {
        const keys = [
            "Ngày",
            "Ngay",
            "Ngày nhập",
            "Ngay nhap",
            "Ngày chứng từ",
            "Ngay chung tu",
            "Ngày giao dịch",
            "Ngay giao dich"
        ];
        for (const k of keys) {
            if (row[k] == null || row[k] === "") continue;
            let d = parseCellToDate(row[k]);
            if (!d || isNaN(d.getTime())) {
                const s = cellDisplayString(row[k]).trim();
                if (s) d = parseDateFlexible(s);
            }
            if (d && !isNaN(d.getTime())) return d;
        }
        return null;
    }

    function nxlcRowInSelectedPeriod(row) {
        const d = pickNxlcNgay(row);
        if (!d || isNaN(d.getTime())) return false;
        const fromRaw = document.getElementById("tkh-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("tkh-filter-to-date")?.value?.trim() ?? "";
        let fromD = fromRaw ? parseDateFlexible(fromRaw) : null;
        let toD = toRaw ? parseDateFlexible(toRaw) : null;
        if (!fromD && toD) fromD = toD;
        if (!fromD) return false;
        if (!toD) toD = fromD;
        const start = new Date(fromD.getFullYear(), fromD.getMonth(), fromD.getDate(), 0, 0, 0, 0);
        const end = new Date(toD.getFullYear(), toD.getMonth(), toD.getDate(), 23, 59, 59, 999);
        return d >= start && d <= end;
    }

    /** Có chọn kỳ (Từ/Đến) thì lọc theo ngày NXLC; không chọn thì lấy toàn bộ phiếu (để cột Nhập/Xuất vẫn có số liệu). */
    function nxlcRowPassesDateFilterIfAny(row) {
        if (!hasTkhNxlcPeriodFilter()) return true;
        return nxlcRowInSelectedPeriod(row);
    }

    /** Có ít nhất một ngày hợp lệ trên Từ/Đến — khi đó mới lọc phiếu NXLC theo kỳ (xem nxlcRowPassesDateFilterIfAny). */
    function hasTkhNxlcPeriodFilter() {
        const fromRaw = document.getElementById("tkh-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("tkh-filter-to-date")?.value?.trim() ?? "";
        const fromD = fromRaw ? parseDateFlexible(fromRaw) : null;
        const toD = toRaw ? parseDateFlexible(toRaw) : null;
        return !!(fromD || toD);
    }

    function pickNxlcTenNhaCungCap(row) {
        return pickFirstCell(row, [
            "Tên nhà cung cấp",
            "Ten nha cung cap",
            "Nhà cung cấp",
            "Nha cung cap",
            "NCC",
            "Tên NCC",
            "Ten NCC"
        ]);
    }

    function pickNxlcTenNoiNhap(row) {
        return pickFirstCell(row, [
            "Tên nơi nhập",
            "Ten noi nhap",
            "Nơi nhập",
            "Noi nhap",
            "Nơi nhập hàng",
            "Kho nhập",
            "Noi nhap hang"
        ]);
    }

    function pickNxlcXeCapDau(row) {
        return pickFirstCell(row, [
            "Xe cấp dầu",
            "Xe cap dau",
            "Xe cấp NL",
            "Xe cap NL",
            "Xe giao dầu",
            "Xe giao dau"
        ]);
    }

    function getNxlcTenNhienLieuCell(r) {
        return pickFirstCell(r, [
            "Tên nhiên liệu",
            "Ten nhien lieu",
            "Tên NL",
            "Ten NL",
            "Loại nhiên liệu",
            "Loai nhien lieu"
        ]);
    }

    function parseNxlcSoLuongNl(row) {
        const n = pickNumericByAliases(
            row,
            [
                "Số lượng NL",
                "So luong NL",
                "SL NL",
                "Sl NL",
                "Số lượng nhiên liệu",
                "So luong nhien lieu",
                "Số lượng",
                "So luong",
                "SL",
                "Sl",
                "Lượng NL",
                "Luong NL"
            ],
            /s[oố]\s*l[ưu]ợ?ng.*nl|so\s*luong.*nl|lu[oọ]ng.*nl/i
        );
        if (n != null) return n;
        for (const [k, v] of Object.entries(row || {})) {
            if (!/s[oố]\s*l|so\s*luong|lu[oọ]ng/i.test(String(k))) continue;
            if (!/nl|nhien\s*lieu|nhi[eê]n/i.test(String(k))) continue;
            const q = parseNumeric(v);
            if (q != null && !isNaN(q)) return q;
        }
        return null;
    }

    function parseNxlcThanhTienNl(row) {
        const n = pickNumericByAliases(
            row,
            [
                "Thành tiền NL",
                "Thanh tien NL",
                "Thành tiền nhiên liệu",
                "Thanh tien nhien lieu",
                "Thành tiền",
                "Thanh tien",
                "Giá trị NL",
                "Gia tri NL",
                "Giá trị nhiên liệu",
                "Gia tri nhien lieu"
            ],
            /th[aà]nh\s*ti[eề]n.*nl|thanh\s*tien.*nl|gia\s*tri.*nl|gia\s*tri.*nhien|thanh\s*tien.*nhien/i
        );
        if (n != null) return n;
        for (const [k, v] of Object.entries(row || {})) {
            const key = String(k);
            if (!/th[aà]nh\s*ti[eề]n|thanh\s*tien|gia\s*tri/i.test(key)) continue;
            if (!/nl|nhien\s*lieu|nhi[eê]n/i.test(key)) continue;
            const q = parseNumeric(v);
            if (q != null && !isNaN(q)) return q;
        }
        return null;
    }

    /**
     * Tổng SL theo khóa normalizeKeyPart(Tên nhiên liệu): phiếu khớp rowPredicate + khoảng ngày + 3 filter.
     */
    function buildNxlcSlAggregateFromNxlc(nxlcRows, rowPredicate) {
        const map = new Map();
        if (!nxlcRows?.length || typeof rowPredicate !== "function") return map;
        for (const r of nxlcRows) {
            if (!rowPredicate(r)) continue;
            if (!nxlcRowPassesDateFilterIfAny(r)) continue;
            const nccSel = document.getElementById("tkh-filter-ncc")?.value?.trim() ?? "";
            const ctSel = document.getElementById("tkh-filter-cong-trinh")?.value?.trim() ?? "";
            const khoSel = document.getElementById("tkh-filter-ten-kho")?.value?.trim() ?? "";
            if (nccSel && pickNxlcTenNhaCungCap(r).trim() !== nccSel) continue;
            if (ctSel && pickNxlcTenNoiNhap(r).trim() !== ctSel) continue;
            if (khoSel && pickNxlcXeCapDau(r).trim() !== khoSel) continue;
            const tenNl = getNxlcTenNhienLieuCell(r).trim();
            if (!tenNl) continue;
            const nk = normalizeKeyPart(tenNl);
            const sl = parseNxlcSoLuongNl(r);
            if (sl == null || isNaN(sl)) continue;
            map.set(nk, (map.get(nk) || 0) + sl);
        }
        return map;
    }

    function buildNxlcGtAggregateFromNxlc(nxlcRows, rowPredicate) {
        const map = new Map();
        if (!nxlcRows?.length || typeof rowPredicate !== "function") return map;
        for (const r of nxlcRows) {
            if (!rowPredicate(r)) continue;
            if (!nxlcRowPassesDateFilterIfAny(r)) continue;
            const nccSel = document.getElementById("tkh-filter-ncc")?.value?.trim() ?? "";
            const ctSel = document.getElementById("tkh-filter-cong-trinh")?.value?.trim() ?? "";
            const khoSel = document.getElementById("tkh-filter-ten-kho")?.value?.trim() ?? "";
            if (nccSel && pickNxlcTenNhaCungCap(r).trim() !== nccSel) continue;
            if (ctSel && pickNxlcTenNoiNhap(r).trim() !== ctSel) continue;
            if (khoSel && pickNxlcXeCapDau(r).trim() !== khoSel) continue;
            const tenNl = getNxlcTenNhienLieuCell(r).trim();
            if (!tenNl) continue;
            const nk = normalizeKeyPart(tenNl);
            const gt = parseNxlcThanhTienNl(r);
            if (gt == null || isNaN(gt)) continue;
            map.set(nk, (map.get(nk) || 0) + gt);
        }
        return map;
    }

    function buildNhapSlAggregateFromNxlc(nxlcRows) {
        return buildNxlcSlAggregateFromNxlc(nxlcRows, nxlcRowIsPhieuNhap);
    }

    function buildXuatSlAggregateFromNxlc(nxlcRows) {
        return buildNxlcSlAggregateFromNxlc(nxlcRows, nxlcRowIsPhieuXuat);
    }

    function buildNhapGtAggregateFromNxlc(nxlcRows) {
        return buildNxlcGtAggregateFromNxlc(nxlcRows, nxlcRowIsPhieuNhap);
    }

    function buildXuatGtAggregateFromNxlc(nxlcRows) {
        return buildNxlcGtAggregateFromNxlc(nxlcRows, nxlcRowIsPhieuXuat);
    }

    /**
     * Theo tên tài sản / tên nhiên liệu trên «Danh sách tài sản», xác định thứ tự khóa
     * normalizeKeyPart(...) tương ứng với cột «Tên nhiên liệu» trên NXLC cần lấy:
     * - Dầu cầu (DS) ← «Dầu cầu»
     * - Mỡ / Mỡ bò (DS) ← ưu tiên «Mỡ», sau đó «Mỡ bò»
     * - Nhớt / Nhớt động cơ (DS) ← ưu tiên «Nhớt», sau đó «Nhớt động cơ»
     * - Dầu thủy lực (DS) ← «Dầu thủy lực»
     * Các tên khác: khớp đúng tên DS trước, rồi fallback mềm.
     */
    function nxlcLookupKeysForDsTenNl(tenDs) {
        const t = String(tenDs ?? "").trim();
        if (!t) return [];
        const nk = normalizeKeyPart(t);

        if (/dầu\s*thủy\s*lực/i.test(t)) {
            return [normalizeKeyPart("Dầu thủy lực")];
        }
        if (/dầu\s*cầu/i.test(t)) {
            return [normalizeKeyPart("Dầu cầu")];
        }
        if (/dầu\s*diese|dầu\s*do\b|diezel|diesel/i.test(t)) {
            const keys = [
                normalizeKeyPart("Dầu Diesel"),
                normalizeKeyPart("Dầu Diezel"),
                normalizeKeyPart("Dầu DO"),
                nk
            ];
            return [...new Set(keys)];
        }
        if (/\bmỡ\b|^mỡ|mỡ\s*bò|mỡ\s|mơ\s*bò/i.test(t) || nk === "mobo" || nk === "mo") {
            return [normalizeKeyPart("Mỡ"), normalizeKeyPart("Mỡ bò")];
        }
        if (/nhớt|nhot/i.test(t)) {
            return [normalizeKeyPart("Nhớt"), normalizeKeyPart("Nhớt động cơ")];
        }

        return [nk];
    }

    /**
     * Lấy một giá trị aggregate theo tên NL trên DS: thử lần lượt các khóa NXLC đã map.
     */
    function getAggValueForDsFromNxlcMap(aggMap, tenDs) {
        if (!aggMap) return null;
        const keys = nxlcLookupKeysForDsTenNl(tenDs);
        for (const k of keys) {
            if (!k) continue;
            if (aggMap.has(k)) {
                const v = aggMap.get(k);
                if (v != null && !isNaN(v)) return v;
            }
        }
        return getAggValueByFuelName(aggMap, normalizeKeyPart(tenDs));
    }

    /**
     * Fallback: khớp chuẩn hóa tuyệt đối, rồi khớp chứa 2 chiều (tên khác).
     */
    function getAggValueByFuelName(aggMap, normalizedFuelName) {
        if (!aggMap || !normalizedFuelName) return null;
        if (aggMap.has(normalizedFuelName)) return aggMap.get(normalizedFuelName);
        let bestVal = null;
        let bestScore = 0;
        for (const [k, v] of aggMap.entries()) {
            if (!k) continue;
            const minLen = Math.min(k.length, normalizedFuelName.length);
            if (minLen < 3) continue;
            if (k.includes(normalizedFuelName) || normalizedFuelName.includes(k)) {
                if (minLen > bestScore) {
                    bestScore = minLen;
                    bestVal = v;
                }
            }
        }
        return bestVal;
    }

    function fillSelectOptions(id, sortedValues) {
        const el = document.getElementById(id);
        if (!el) return;
        const prev = el.value;
        el.innerHTML = "";
        const o0 = document.createElement("option");
        o0.value = "";
        o0.textContent = "-- Tất cả --";
        el.appendChild(o0);
        for (const v of sortedValues) {
            const o = document.createElement("option");
            o.value = v;
            o.textContent = v;
            el.appendChild(o);
        }
        if (prev && [...el.options].some((op) => op.value === prev)) el.value = prev;
    }

    function populateTkhNxlcFilters(nxlcRows) {
        const ncc = new Set();
        const ct = new Set();
        const kho = new Set();
        for (const r of nxlcRows || []) {
            const a = pickNxlcTenNhaCungCap(r).trim();
            const b = pickNxlcTenNoiNhap(r).trim();
            const c = pickNxlcXeCapDau(r).trim();
            if (a) ncc.add(a);
            if (b) ct.add(b);
            if (c) kho.add(c);
        }
        const sortVi = (a, b) => a.localeCompare(b, "vi");
        fillSelectOptions("tkh-filter-ncc", [...ncc].sort(sortVi));
        fillSelectOptions("tkh-filter-cong-trinh", [...ct].sort(sortVi));
        fillSelectOptions("tkh-filter-ten-kho", [...kho].sort(sortVi));
    }

    function displayCell(n) {
        if (n == null || isNaN(n)) return "";
        return formatNum(n);
    }

    function computeClosingValue(opening, incoming, outgoing) {
        const hasOpening = opening != null && !isNaN(opening);
        const hasIncoming = incoming != null && !isNaN(incoming);
        const hasOutgoing = outgoing != null && !isNaN(outgoing);
        if (!hasOpening && !hasIncoming && !hasOutgoing) return null;
        const o = hasOpening ? opening : 0;
        const i = hasIncoming ? incoming : 0;
        const x = hasOutgoing ? outgoing : 0;
        return o + i - x;
    }

    /** SL kèm đơn vị nhiên liệu từ DS (cột ĐVT) — vd. `1.234 (Lít)`. */
    function displaySlWithDvtNl(n, dvtRaw) {
        const num = displayCell(n);
        if (num === "") return "";
        const u = String(dvtRaw ?? "")
            .trim()
            .replace(/\s+/g, " ");
        if (!u) return num;
        return `${num} (${u})`;
    }

    function renderTable(rows) {
        const tbody = document.getElementById("tkh-tbody");
        if (!tbody) return;

        const totals = {
            tonDauSl: 0,
            tonDauGt: 0,
            nhapSl: 0,
            nhapGt: 0,
            xuatSl: 0,
            xuatGt: 0,
            tonCuoiSl: 0,
            tonCuoiGt: 0
        };
        let anyNum = false;

        function addTot(st) {
            for (const k of Object.keys(totals)) {
                const v = st[k];
                if (v != null && !isNaN(v)) {
                    totals[k] += v;
                    anyNum = true;
                }
            }
        }

        const selectedDateRange = getSelectedDateRangeFromDom();
        const tonDkAgg = buildTonKhoDkAggregateByTenNl(cacheTonKhoDkRows, selectedDateRange);

        const useNxlcSlKy = Array.isArray(cacheNxlcCtRows) && !lastNxlcFetchError;
        const nhapSlAgg = useNxlcSlKy ? buildNhapSlAggregateFromNxlc(cacheNxlcCtRows) : null;
        const xuatSlAgg = useNxlcSlKy ? buildXuatSlAggregateFromNxlc(cacheNxlcCtRows) : null;
        const nhapGtAgg = useNxlcSlKy ? buildNhapGtAggregateFromNxlc(cacheNxlcCtRows) : null;
        const xuatGtAgg = useNxlcSlKy ? buildXuatGtAggregateFromNxlc(cacheNxlcCtRows) : null;

        let html = "";
        const sorted = [...(rows || [])].sort((a, b) => {
            const ma = pickDsMaThietBi(a).localeCompare(pickDsMaThietBi(b), "vi");
            if (ma !== 0) return ma;
            return pickDsTenNhienLieu(a).localeCompare(pickDsTenNhienLieu(b), "vi");
        });

        for (const r of sorted) {
            const stBase = extractStockCells(r);
            const tenNl = pickDsTenNhienLieu(r);
            const nk = normalizeKeyPart(tenNl);
            let tonDauSl = null;
            let tonDauGt = stBase.tonDauGt;
            const hit = tonDkAgg.get(nk);
            if (hit) {
                tonDauSl = hit.sl;
                if (hit.gt != null && !isNaN(hit.gt)) tonDauGt = hit.gt;
            }
            let nhapSl = stBase.nhapSl;
            let nhapGt = stBase.nhapGt;
            if (useNxlcSlKy && nhapSlAgg) {
                const v = getAggValueForDsFromNxlcMap(nhapSlAgg, tenNl);
                nhapSl = v != null && !isNaN(v) ? v : 0;
            }
            if (useNxlcSlKy && nhapGtAgg) {
                const v = getAggValueForDsFromNxlcMap(nhapGtAgg, tenNl);
                nhapGt = v != null && !isNaN(v) ? v : 0;
            }
            let xuatSl = stBase.xuatSl;
            let xuatGt = stBase.xuatGt;
            if (useNxlcSlKy && xuatSlAgg) {
                const v = getAggValueForDsFromNxlcMap(xuatSlAgg, tenNl);
                xuatSl = v != null && !isNaN(v) ? v : 0;
            }
            if (useNxlcSlKy && xuatGtAgg) {
                const v = getAggValueForDsFromNxlcMap(xuatGtAgg, tenNl);
                xuatGt = v != null && !isNaN(v) ? v : 0;
            }
            const tonCuoiSl = computeClosingValue(tonDauSl, nhapSl, xuatSl);
            const tonCuoiGt = computeClosingValue(tonDauGt, nhapGt, xuatGt);
            const st = { ...stBase, tonDauSl, tonDauGt, nhapSl, nhapGt, xuatSl, xuatGt, tonCuoiSl, tonCuoiGt };
            addTot(st);
            const maNl = pickDsMaThietBi(r);
            const dvt = pickDsDvt(r);
            const km = pickDsKhoanMucChiPhi(r);
            const gc = pickDsGhiChu(r);
            html += `<tr class="tkh-data-row">
                <td>${escapeHtml(maNl)}</td>
                <td style="text-align:left;">${escapeHtml(tenNl)}</td>
                <td>${escapeHtml(dvt)}</td>
                <td>${escapeHtml(displaySlWithDvtNl(st.tonDauSl, dvt))}</td>
                <td class="tkh-val-cell">${escapeHtml(displayCell(st.tonDauGt))}</td>
                <td>${escapeHtml(displaySlWithDvtNl(st.nhapSl, dvt))}</td>
                <td class="tkh-val-cell">${escapeHtml(displayCell(st.nhapGt))}</td>
                <td>${escapeHtml(displaySlWithDvtNl(st.xuatSl, dvt))}</td>
                <td class="tkh-val-cell">${escapeHtml(displayCell(st.xuatGt))}</td>
                <td>${escapeHtml(displaySlWithDvtNl(st.tonCuoiSl, dvt))}</td>
                <td class="tkh-val-cell">${escapeHtml(displayCell(st.tonCuoiGt))}</td>
                <td style="text-align:left;">${escapeHtml(km)}</td>
                <td style="text-align:left;">${escapeHtml(gc)}</td>
            </tr>`;
        }

        if (!sorted.length) {
            html =
                '<tr><td colspan="13" style="padding:12px;">Không có dòng «Danh sách tài sản» nào có «Tên nhóm» = «Nhiên liệu».</td></tr>';
        }

        const tf = (k) => (anyNum ? formatNum(totals[k]) : "");
        html += `<tr class="inventory-total-row">
            <td></td>
            <td>TỔNG CỘNG</td>
            <td></td>
            <td>${escapeHtml(tf("tonDauSl"))}</td>
            <td class="tkh-val-cell">${escapeHtml(tf("tonDauGt"))}</td>
            <td>${escapeHtml(tf("nhapSl"))}</td>
            <td class="tkh-val-cell">${escapeHtml(tf("nhapGt"))}</td>
            <td>${escapeHtml(tf("xuatSl"))}</td>
            <td class="tkh-val-cell">${escapeHtml(tf("xuatGt"))}</td>
            <td>${escapeHtml(tf("tonCuoiSl"))}</td>
            <td class="tkh-val-cell">${escapeHtml(tf("tonCuoiGt"))}</td>
            <td colspan="2"></td>
        </tr>`;

        tbody.innerHTML = html;
    }

    function setStatus(msg, isErr) {
        const el = document.getElementById("tkh-appsheet-status");
        if (!el) return;
        el.textContent = msg;
        el.classList.toggle("text-red-600", !!isErr);
        el.classList.toggle("text-gray-700", !isErr);
    }

    const TKH_PREVIEW_ROW_CAP = 500;

    /** Gom mọi tên cột từ toàn bộ dòng (tránh sót cột chỉ có ở vài dòng). */
    function collectKeysFromRows(rows) {
        const set = new Set();
        if (!rows?.length) return [];
        const n = Math.min(rows.length, 20000);
        for (let i = 0; i < n; i++) {
            const r = rows[i];
            if (r && typeof r === "object") Object.keys(r).forEach((k) => set.add(k));
        }
        return orderPreviewColumnKeys([...set]);
    }

    /**
     * Đưa cột ngày/giờ lên trước (dễ thấy); còn lại sắp theo vi.
     * Trước đây sort A→Z khiến «Ngày» nằm giữa/cuối bảng rộng, dễ tưởng «thiếu cột».
     */
    function orderPreviewColumnKeys(keys) {
        const all = keys.filter(Boolean);
        const pri = all.filter(isLikelyDateTimeColumnKey).sort((a, b) => a.localeCompare(b, "vi"));
        const rest = all.filter((k) => !isLikelyDateTimeColumnKey(k)).sort((a, b) => a.localeCompare(b, "vi"));
        return [...pri, ...rest];
    }

    function previewCellValueForModal(v, columnKey) {
        if (v == null) return "";
        const col = columnKey != null ? String(columnKey) : "";
        const rawDisp = cellDisplayString(v).trim();

        const shouldFormatDate =
            (col && isLikelyDateTimeColumnKey(col)) || (rawDisp && looksLikeRawDateString(rawDisp));

        if (shouldFormatDate) {
            const dd = cellValueToDdMmYyyyStr(v);
            if (dd) return dd.length > 800 ? `${dd.slice(0, 800)}…` : dd;
        }

        let s = rawDisp;
        if (!s && typeof v === "object" && v !== null) {
            const raw = v.Value ?? v.value ?? v.DisplayValue ?? v.displayValue ?? v.Text ?? v.text;
            if (raw != null) s = String(raw).trim();
        }
        if (!s && typeof v === "object" && v !== null) {
            try {
                const j = JSON.stringify(v);
                if (j && j !== "{}" && j !== "null") s = j;
            } catch (_) {
                s = "[object]";
            }
        }
        if (!s) s = String(v);
        return s.length > 800 ? `${s.slice(0, 800)}…` : s;
    }

    function buildRawPreviewTableHtml(rows) {
        const r = rows ?? [];
        if (!r.length) {
            return '<p class="tkh-prev-note">Không có dòng.</p>';
        }
        const total = r.length;
        const show = r.slice(0, TKH_PREVIEW_ROW_CAP);
        const keys = collectKeysFromRows(r);
        if (!keys.length) {
            return `<p class="tkh-prev-note">${total} dòng — không có khóa cột.</p>`;
        }
        const thead = `<tr>${keys.map((k) => `<th>${escapeHtml(k)}</th>`).join("")}</tr>`;
        const tbody = show
            .map(
                (row) =>
                    `<tr>${keys
                        .map((k) => `<td>${escapeHtml(previewCellValueForModal(row[k], k))}</td>`)
                        .join("")}</tr>`
            )
            .join("");
        const note =
            (total > TKH_PREVIEW_ROW_CAP
                ? `Hiển thị tối đa ${TKH_PREVIEW_ROW_CAP} / ${total} dòng. `
                : `${total} dòng. `) +
            "Cột ngày/giờ xếp trước; giá trị ngày hiển thị dd/mm/yyyy. Kéo ngang nếu bảng rộng.";
        return `<p class="tkh-prev-note">${escapeHtml(note)}</p><div class="tkh-prev-scroll"><table class="tkh-prev-table"><thead>${thead}</thead><tbody>${tbody}</tbody></table></div>`;
    }

    function switchTkhLoadedTab(which) {
        document.querySelectorAll(".tkh-loaded-tab").forEach((b) => {
            b.classList.toggle("active", b.getAttribute("data-tkh-prev") === which);
        });
        document.querySelectorAll(".tkh-loaded-panel").forEach((p) => {
            p.classList.toggle("is-visible", p.id === `tkh-loaded-panel-${which}`);
        });
    }

    function rowDateInRange(dateValue, dateRange) {
        if (!dateRange) return true;
        if (!dateValue || isNaN(dateValue.getTime())) return false;
        return dateValue >= dateRange.start && dateValue <= dateRange.end;
    }

    function formatPreviewTabCount(filteredCount, totalCount) {
        if (totalCount === filteredCount) return String(totalCount);
        return `${filteredCount}/${totalCount}`;
    }

    function rowMatchesNxlcCurrentFilters(row, dateRange, nccSel, ctSel, khoSel) {
        if (dateRange) {
            const d = pickNxlcNgay(row);
            if (!rowDateInRange(d, dateRange)) return false;
        }
        if (nccSel && pickNxlcTenNhaCungCap(row).trim() !== nccSel) return false;
        if (ctSel && pickNxlcTenNoiNhap(row).trim() !== ctSel) return false;
        if (khoSel && pickNxlcXeCapDau(row).trim() !== khoSel) return false;
        return true;
    }

    function openTkhLoadedDataModal() {
        const modal = document.getElementById("tkh-loaded-data-modal");
        if (!modal) return;
        if (cacheDsRows == null) {
            alert("Chưa tải «Danh sách tài sản». Bấm «Tải dữ liệu AppSheet» trước.");
            return;
        }
        const dateRange = getSelectedDateRangeFromDom();
        const nccSel = document.getElementById("tkh-filter-ncc")?.value?.trim() ?? "";
        const ctSel = document.getElementById("tkh-filter-cong-trinh")?.value?.trim() ?? "";
        const khoSel = document.getElementById("tkh-filter-ten-kho")?.value?.trim() ?? "";

        const dsRowsAll = cacheDsRows ?? [];
        const dsRowsView = dsRowsAll.filter(rowIsNhomNhienLieu);

        const dkRowsAll = cacheTonKhoDkRows ?? [];
        const dkRowsView = dkRowsAll.filter((r) => {
            if (!dateRange) return true;
            const rawNgay = pickTkNgayNhapSoDu(r);
            const d = rawNgay != null ? parseCellToDate(rawNgay) : null;
            return rowDateInRange(d, dateRange);
        });

        const nxRowsAll = cacheNxlcCtRows ?? [];
        const nxRowsView = nxRowsAll.filter((r) => rowMatchesNxlcCurrentFilters(r, dateRange, nccSel, ctSel, khoSel));

        const dsEl = document.getElementById("tkh-loaded-panel-ds");
        const dkEl = document.getElementById("tkh-loaded-panel-dk");
        const nxEl = document.getElementById("tkh-loaded-panel-nxlc");
        if (dsEl) dsEl.innerHTML = buildRawPreviewTableHtml(dsRowsView);
        if (dkEl) dkEl.innerHTML = buildRawPreviewTableHtml(dkRowsView);
        if (nxEl) {
            let inner = buildRawPreviewTableHtml(nxRowsView);
            if (lastNxlcFetchError) {
                inner =
                    `<p class="tkh-prev-note" style="color:#b91c1c;font-weight:600;">«${TABLE_NXLC_CT}»: ${escapeHtml(lastNxlcFetchError)}</p>` +
                    inner;
            }
            nxEl.innerHTML = inner;
        }
        const nDs = formatPreviewTabCount(dsRowsView.length, dsRowsAll.length);
        const nDk = formatPreviewTabCount(dkRowsView.length, dkRowsAll.length);
        const nNx = lastNxlcFetchError ? "lỗi" : formatPreviewTabCount(nxRowsView.length, nxRowsAll.length);
        const tDs = document.getElementById("tkh-tab-ds");
        const tDk = document.getElementById("tkh-tab-dk");
        const tNx = document.getElementById("tkh-tab-nxlc");
        if (tDs) tDs.textContent = `Danh sách tài sản (${nDs})`;
        if (tDk) tDk.textContent = `Tồn kho ĐK (${nDk})`;
        if (tNx) tNx.textContent = `Nhập xuất luân chuyển CT (${nNx})`;
        switchTkhLoadedTab("ds");
        modal.style.display = "flex";
        modal.setAttribute("aria-hidden", "false");
        document.body.style.overflow = "hidden";
    }

    function closeTkhLoadedDataModal() {
        const modal = document.getElementById("tkh-loaded-data-modal");
        if (!modal) return;
        modal.style.display = "none";
        modal.setAttribute("aria-hidden", "true");
        document.body.style.overflow = "";
    }

    async function loadFromAppSheet(forceRefresh) {
        const btn = document.getElementById("tkh-btn-load");
        const refBtn = document.getElementById("tkh-btn-refresh");
        if (btn) btn.disabled = true;
        if (refBtn) refBtn.disabled = true;
        setStatus("Đang tải AppSheet…", false);
        try {
            if (forceRefresh) {
                cacheDsRows = null;
                cacheTonKhoDkRows = null;
                cacheNxlcCtRows = null;
                lastNxlcFetchError = null;
            }
            if (!cacheDsRows) {
                setStatus(`Đang tải «${TABLE_DS_TAI_SAN}»…`, false);
                cacheDsRows = await fetchAppSheetTable(TABLE_DS_TAI_SAN);
            }
            if (!cacheTonKhoDkRows) {
                setStatus(`Đang tải «${TABLE_TON_KHO_DK}»…`, false);
                cacheTonKhoDkRows = await fetchAppSheetTable(TABLE_TON_KHO_DK);
            }
            let nxlcLoadWarn = null;
            if (cacheNxlcCtRows == null) {
                setStatus(`Đang tải «${TABLE_NXLC_CT}»…`, false);
                try {
                    cacheNxlcCtRows = await fetchAppSheetTable(TABLE_NXLC_CT);
                    lastNxlcFetchError = null;
                } catch (nxErr) {
                    console.error(nxErr);
                    cacheNxlcCtRows = [];
                    lastNxlcFetchError = nxErr.message || String(nxErr);
                    nxlcLoadWarn = lastNxlcFetchError;
                }
            }
            populateTkhNxlcFilters(cacheNxlcCtRows || []);
            syncTkhBanners();
            const fuelRows = (cacheDsRows || []).filter(rowIsNhomNhienLieu);
            cacheFuelRows = fuelRows;
            renderTable(cacheFuelRows);
            const n = cacheDsRows?.length ?? 0;
            const m = fuelRows.length;
            const t = cacheTonKhoDkRows?.length ?? 0;
            const x = cacheNxlcCtRows?.length ?? 0;
            let statusMsg = `Đã tải ${n} dòng «${TABLE_DS_TAI_SAN}», ${t} dòng «${TABLE_TON_KHO_DK}», ${x} dòng «${TABLE_NXLC_CT}» — ${m} dòng nhóm «${NHOM_NHIEN_LIEU_LABEL}».`;
            if (nxlcLoadWarn) {
                statusMsg += ` — không tải NXLC: ${nxlcLoadWarn} (nhập kỳ: DS).`;
            }
            setStatus(statusMsg, !!nxlcLoadWarn);
        } catch (e) {
            console.error(e);
            setStatus(`Lỗi: ${e.message}`, true);
            cacheFuelRows = [];
            renderTable([]);
        } finally {
            if (btn) btn.disabled = false;
            if (refBtn) refBtn.disabled = false;
        }
    }

    function init() {
        const loadBtn = document.getElementById("tkh-btn-load");
        const refBtn = document.getElementById("tkh-btn-refresh");
        if (loadBtn) loadBtn.addEventListener("click", () => loadFromAppSheet(false));
        if (refBtn) refBtn.addEventListener("click", () => loadFromAppSheet(true));

        const search = document.getElementById("tkh-search");
        if (search) {
            search.addEventListener("input", () => {
                const q = search.value.trim().toLowerCase();
                const tbody = document.getElementById("tkh-tbody");
                if (!tbody) return;
                for (const tr of tbody.querySelectorAll("tr.tkh-data-row")) {
                    tr.style.display = !q || tr.textContent.toLowerCase().includes(q) ? "" : "none";
                }
            });
        }

        const rerenderTableView = () => renderTable(cacheFuelRows);
        for (const id of ["tkh-filter-from-date", "tkh-filter-to-date"]) {
            const el = document.getElementById(id);
            if (!el) continue;
            el.addEventListener("change", rerenderTableView);
            el.addEventListener("input", rerenderTableView);
        }
        for (const id of ["tkh-filter-ncc", "tkh-filter-cong-trinh", "tkh-filter-ten-kho"]) {
            document.getElementById(id)?.addEventListener("change", () => {
                if (id === "tkh-filter-cong-trinh") syncTkhCongTrinhBanner();
                rerenderTableView();
            });
        }

        wireNativeDatePickerButton("tkh-filter-from-date", "tkh-filter-from-date-picker", "tkh-filter-from-date-cal");
        wireNativeDatePickerButton("tkh-filter-to-date", "tkh-filter-to-date-picker", "tkh-filter-to-date-cal");

        for (const id of ["tkh-filter-from-date", "tkh-filter-to-date"]) {
            const el = document.getElementById(id);
            if (!el) continue;
            const onDateFilterChange = () => syncTkhBanners();
            el.addEventListener("input", onDateFilterChange);
            el.addEventListener("keyup", onDateFilterChange);
            el.addEventListener("change", onDateFilterChange);
            el.addEventListener("blur", onDateFilterChange);
            el.addEventListener("paste", () => setTimeout(onDateFilterChange, 0));
        }
        syncTkhBanners();

        document.getElementById("tkh-btn-view-loaded-tables")?.addEventListener("click", openTkhLoadedDataModal);
        document.getElementById("tkh-loaded-close")?.addEventListener("click", closeTkhLoadedDataModal);
        document.getElementById("tkh-loaded-data-modal")?.addEventListener("click", (e) => {
            if (e.target && e.target.id === "tkh-loaded-data-modal") closeTkhLoadedDataModal();
        });
        document.querySelectorAll(".tkh-loaded-tab").forEach((btn) => {
            btn.addEventListener("click", () => switchTkhLoadedTab(btn.getAttribute("data-tkh-prev") || "ds"));
        });
        document.addEventListener("keydown", (e) => {
            if (e.key !== "Escape") return;
            const modal = document.getElementById("tkh-loaded-data-modal");
            if (modal && modal.style.display === "flex") closeTkhLoadedDataModal();
        });

        loadFromAppSheet(false);
    }

    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
    else init();
})();
