/**
 * Báo cáo tổng hợp xuất nhiên liệu theo công trình — AppSheet.
 * «Nhập trong tháng»; xuất theo công trình: ưu «Tên nơi nhập», trống thì «Tên nơi xuất»; «Tổng xuất», «Tồn cuối tháng» (SL), «Ghi chú»
 * lấy từ «Nhập xuất luân chuyển CT» (cột «Loại phiếu», «Số lượng NL», «Ghi chú»…).
 * «Tồn đầu tháng» (SL): chỉ từ «Tồn kho ĐK» — cột «Sl tồn ĐK», khớp «Tên NL» với tên nhiên liệu,
 * «Ngày nhập số dư ĐK» nằm trong khoảng «Từ ngày»..«Đến ngày» (cùng báo cáo tồn kho NL).
 * «Tồn cuối tháng» = Tồn đầu + Nhập trong tháng − Tổng xuất (cùng kỳ), các thành phần SL từ NXLC + tồn đầu từ ĐK.
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
    let cacheNxlcCtRows = null;
    let lastNxlcFetchError = null;
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

    function pickDsTenNhom(row) {
        const tn = pickFirstCell(row, ["Tên nhóm", "Ten nhom", "TenNhom"]);
        if (tn) return tn;
        return pickFirstCell(row, [
            "Nhóm thiết bị",
            "Nhom thiet bi",
            "Nhóm xe máy thiết bị",
            "Nhom xe may thiet bi"
        ]);
    }

    /** Bộ lọc «Nhóm nhiên liệu»: ưu tiên cột phụ trên DS, không có thì «Nhóm thiết bị». */
    function pickDsNhomNhienLieuPhu(row) {
        const direct = pickFirstCell(row, [
            "Nhóm nhiên liệu",
            "Nhom nhien lieu",
            "Loại nhiên liệu",
            "Loai nhien lieu",
            "Phân loại nhiên liệu",
            "Phan loai nhien lieu"
        ]);
        if (direct) return direct;
        return pickFirstCell(row, ["Nhóm thiết bị", "Nhom thiet bi", "Nhóm xe máy thiết bị", "Nhom xe may thiet bi"]);
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

    const THXL_DEFAULT_BANNER_SITE_LINE = "CT HỐ SÓI - SÔNG ĐUỐNG";

    function syncThxlDateRangeBanner() {
        const dr = document.getElementById("thxl-date-range");
        if (!dr) return;
        const fromRaw = document.getElementById("thxl-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("thxl-filter-to-date")?.value?.trim() ?? "";
        const dFrom = fromRaw ? parseDateFlexible(fromRaw) : null;
        const dTo = toRaw ? parseDateFlexible(toRaw) : null;
        const labelFrom =
            !fromRaw ? "…" : dFrom && !isNaN(dFrom.getTime()) ? formatDateVn(dFrom) : fromRaw;
        const labelTo = !toRaw ? "…" : dTo && !isNaN(dTo.getTime()) ? formatDateVn(dTo) : toRaw;
        dr.textContent = `(Từ ngày ${labelFrom} tới ngày ${labelTo})`;
    }

    function syncThxlCongTrinhBanner() {
        const el = document.getElementById("thxl-banner-cong-trinh-line");
        const sel = document.getElementById("thxl-filter-cong-trinh");
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
            el.textContent = THXL_DEFAULT_BANNER_SITE_LINE;
            return;
        }
        if (options.length === 1) {
            el.textContent = options[0];
            return;
        }
        el.textContent = options.join(" • ");
    }

    function syncThxlBanners() {
        syncThxlDateRangeBanner();
        syncThxlCongTrinhBanner();
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
            if (a > 12 && b <= 12) return buildStrictDate(y, b, a);
            if (a <= 12 && b > 12) return buildStrictDate(y, a, b);
            return buildStrictDate(y, a, b);
        }
        const d = new Date(s);
        if (!isNaN(d.getTime())) return d;
        return null;
    }

    function getSelectedDateRangeFromDom() {
        const fromRaw = document.getElementById("thxl-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("thxl-filter-to-date")?.value?.trim() ?? "";
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
            const prev = map.get(nk) || { slSum: 0, hasSl: false };
            if (sl != null && !isNaN(sl)) {
                prev.slSum += sl;
                prev.hasSl = true;
            }
            map.set(nk, prev);
        }
        const out = new Map();
        for (const [k, v] of map) {
            out.set(k, { sl: v.hasSl ? v.slSum : null });
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

    function nxlcRowIsPhieuXuat(r) {
        const lp = getLoaiPhieuColumnNxlc(r);
        const sLoai = cellDisplayString(lp).trim();
        if (!sLoai) return false;
        if (sLoai === "Nhập" || sLoai === "NHẬP" || loaiPhieuIsNhap(sLoai)) return false;
        if (sLoai === "Xuất" || sLoai === "XUẤT") return true;
        return loaiPhieuIsXuat(sLoai);
    }

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
        const fromRaw = document.getElementById("thxl-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("thxl-filter-to-date")?.value?.trim() ?? "";
        let fromD = fromRaw ? parseDateFlexible(fromRaw) : null;
        let toD = toRaw ? parseDateFlexible(toRaw) : null;
        if (!fromD && toD) fromD = toD;
        if (!fromD) return false;
        if (!toD) toD = fromD;
        const start = new Date(fromD.getFullYear(), fromD.getMonth(), fromD.getDate(), 0, 0, 0, 0);
        const end = new Date(toD.getFullYear(), toD.getMonth(), toD.getDate(), 23, 59, 59, 999);
        return d >= start && d <= end;
    }

    function hasThxlNxlcPeriodFilter() {
        const fromRaw = document.getElementById("thxl-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("thxl-filter-to-date")?.value?.trim() ?? "";
        const fromD = fromRaw ? parseDateFlexible(fromRaw) : null;
        const toD = toRaw ? parseDateFlexible(toRaw) : null;
        return !!(fromD || toD);
    }

    function nxlcRowPassesDateFilterIfAny(row) {
        if (!hasThxlNxlcPeriodFilter()) return true;
        return nxlcRowInSelectedPeriod(row);
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

    function pickNxlcTenNoiXuat(row) {
        return pickFirstCell(row, [
            "Tên nơi xuất",
            "Ten noi xuat",
            "Nơi xuất",
            "Noi xuat",
            "Nơi xuất hàng",
            "Kho xuất",
            "Noi xuat hang"
        ]);
    }

    /**
     * Cột xuất theo công trình: ưu «Tên nơi nhập», không có thì «Tên nơi xuất», còn trống mới gộp «(Chưa ghi…)».
     */
    function pickNxlcDestinationForXuatColumn(row) {
        const n = pickNxlcTenNoiNhap(row).trim();
        if (n) return n;
        const x = pickNxlcTenNoiXuat(row).trim();
        if (x) return x;
        return "(Chưa ghi nơi nhập / nơi xuất)";
    }

    /** Khớp bộ lọc «Công trình»: đúng «Tên nơi nhập» hoặc «Tên nơi xuất». */
    function nxlcRowMatchesCongTrinhFilter(row, ctSel) {
        if (!ctSel) return true;
        if (pickNxlcTenNoiNhap(row).trim() === ctSel) return true;
        if (pickNxlcTenNoiXuat(row).trim() === ctSel) return true;
        return false;
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

    function pickNxlcGhiChu(row) {
        return pickFirstCell(row, ["Ghi chú", "Ghi chu", "GhiChu", "Mô tả", "Mo ta"]);
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

    function nxlcRowPassesCommonFilters(row) {
        const nccSel = document.getElementById("thxl-filter-ncc")?.value?.trim() ?? "";
        const ctSel = document.getElementById("thxl-filter-cong-trinh")?.value?.trim() ?? "";
        const khoSel = document.getElementById("thxl-filter-ten-kho")?.value?.trim() ?? "";
        if (nccSel && pickNxlcTenNhaCungCap(row).trim() !== nccSel) return false;
        if (ctSel && !nxlcRowMatchesCongTrinhFilter(row, ctSel)) return false;
        if (khoSel && pickNxlcXeCapDau(row).trim() !== khoSel) return false;
        return true;
    }

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

    function mergeGhiChuSetsForDs(ghiMap, tenDs) {
        const keys = nxlcLookupKeysForDsTenNl(tenDs);
        const nkDs = normalizeKeyPart(tenDs);
        const seen = new Set();
        const parts = [];
        function addFromKey(k) {
            if (!k || !ghiMap.has(k)) return;
            for (const g of ghiMap.get(k)) {
                const t = String(g).trim();
                if (!t || seen.has(t)) continue;
                seen.add(t);
                parts.push(t);
            }
        }
        for (const k of keys) addFromKey(k);
        addFromKey(nkDs);
        if (!parts.length) {
            for (const [k, set] of ghiMap.entries()) {
                if (!k || !nkDs || k.length < 3 || nkDs.length < 3) continue;
                if (k.includes(nkDs) || nkDs.includes(k)) {
                    for (const g of set) {
                        const t = String(g).trim();
                        if (!t || seen.has(t)) continue;
                        seen.add(t);
                        parts.push(t);
                    }
                }
            }
        }
        let s = parts.join("; ");
        if (s.length > 400) s = `${s.slice(0, 397)}…`;
        return s;
    }

    function buildNxlcSlAggregateFromNxlc(nxlcRows, rowPredicate) {
        const map = new Map();
        if (!nxlcRows?.length || typeof rowPredicate !== "function") return map;
        for (const r of nxlcRows) {
            if (!rowPredicate(r)) continue;
            if (!nxlcRowPassesDateFilterIfAny(r)) continue;
            if (!nxlcRowPassesCommonFilters(r)) continue;
            const tenNl = getNxlcTenNhienLieuCell(r).trim();
            if (!tenNl) continue;
            const nk = normalizeKeyPart(tenNl);
            const sl = parseNxlcSoLuongNl(r);
            if (sl == null || isNaN(sl)) continue;
            map.set(nk, (map.get(nk) || 0) + sl);
        }
        return map;
    }

    /**
     * Xuất theo công trình: cột = «Tên nơi nhập», nếu trống thì «Tên nơi xuất».
     * Map(destLabel -> Map(nk -> tổng SL)).
     */
    function buildXuatSlByDestinationFromNxlc(nxlcRows) {
        const out = new Map();
        if (!nxlcRows?.length) return out;
        for (const r of nxlcRows) {
            if (!nxlcRowIsPhieuXuat(r)) continue;
            if (!nxlcRowPassesDateFilterIfAny(r)) continue;
            if (!nxlcRowPassesCommonFilters(r)) continue;
            const dest = pickNxlcDestinationForXuatColumn(r);
            const tenNl = getNxlcTenNhienLieuCell(r).trim();
            if (!tenNl) continue;
            const nk = normalizeKeyPart(tenNl);
            const sl = parseNxlcSoLuongNl(r);
            if (sl == null || isNaN(sl)) continue;
            if (!out.has(dest)) out.set(dest, new Map());
            const m = out.get(dest);
            m.set(nk, (m.get(nk) || 0) + sl);
        }
        return out;
    }

    function buildGhiChuMapFromNxlc(nxlcRows) {
        const map = new Map();
        if (!nxlcRows?.length) return map;
        for (const r of nxlcRows) {
            if (!nxlcRowPassesDateFilterIfAny(r)) continue;
            if (!nxlcRowPassesCommonFilters(r)) continue;
            const gc = pickNxlcGhiChu(r).trim();
            if (!gc) continue;
            const tenNl = getNxlcTenNhienLieuCell(r).trim();
            if (!tenNl) continue;
            const nk = normalizeKeyPart(tenNl);
            if (!map.has(nk)) map.set(nk, new Set());
            map.get(nk).add(gc);
        }
        return map;
    }

    function getXuatAtDest(xuatByDest, destLabel, tenDs) {
        const fm = xuatByDest.get(destLabel);
        if (!fm) return null;
        return getAggValueForDsFromNxlcMap(fm, tenDs);
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

    function displaySlWithDvtNl(n, dvtRaw) {
        const num = n != null && !isNaN(n) ? formatNum(n) : "";
        if (num === "") return "";
        const u = String(dvtRaw ?? "")
            .trim()
            .replace(/\s+/g, " ");
        if (!u) return num;
        return `${num} (${u})`;
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

    function populateThxlNxlcFilters(nxlcRows) {
        const ncc = new Set();
        const ct = new Set();
        const kho = new Set();
        for (const r of nxlcRows || []) {
            const a = pickNxlcTenNhaCungCap(r).trim();
            const b = pickNxlcTenNoiNhap(r).trim();
            const bx = pickNxlcTenNoiXuat(r).trim();
            const c = pickNxlcXeCapDau(r).trim();
            if (a) ncc.add(a);
            if (b) ct.add(b);
            if (bx) ct.add(bx);
            if (c) kho.add(c);
        }
        const sortVi = (a, b) => a.localeCompare(b, "vi");
        fillSelectOptions("thxl-filter-ncc", [...ncc].sort(sortVi));
        fillSelectOptions("thxl-filter-cong-trinh", [...ct].sort(sortVi));
        fillSelectOptions("thxl-filter-ten-kho", [...kho].sort(sortVi));
    }

    function populateThxlNhomNlFilter(fuelRows) {
        const set = new Set();
        for (const r of fuelRows || []) {
            const v = pickDsNhomNhienLieuPhu(r).trim();
            if (v) set.add(v);
        }
        fillSelectOptions("thxl-filter-nhom-nl", [...set].sort((a, b) => a.localeCompare(b, "vi")));
    }

    function collectXuatDestinationsSorted(xuatByDest) {
        const arr = [...xuatByDest.keys()];
        return arr.sort((a, b) => a.localeCompare(b, "vi"));
    }

    function renderThead(destinations) {
        const thead = document.getElementById("thxl-thead");
        if (!thead) return;
        const n = destinations.length;
        const colSpanXuat = Math.max(1, n);
        const destRow = destinations.length
            ? destinations.map((d) => `<th style="min-width:72px;">${escapeHtml(d)}</th>`).join("")
            : `<th style="min-width:120px;">${escapeHtml("(Không có phiếu xuất theo bộ lọc)")}</th>`;

        thead.innerHTML = `
            <tr>
                <th rowspan="2" style="width:3%;">TT</th>
                <th rowspan="2" style="width:10%;">Tên nhiên liệu</th>
                <th rowspan="2" style="width:4%;">ĐVT</th>
                <th rowspan="2" style="width:6%;" title="Từ bảng «Tồn kho ĐK»: cộng «Sl tồn ĐK» theo «Tên NL», ngày số dư trong khoảng Từ–Đến ngày (giống báo cáo tồn kho NL).">Tồn đầu tháng</th>
                <th rowspan="2" style="width:6%;">Nhập trong tháng</th>
                <th colspan="${colSpanXuat}" style="width:35%;" title="Mỗi cột: «Tên nơi nhập» trên phiếu xuất; nếu trống thì lấy «Tên nơi xuất».">Xuất</th>
                <th rowspan="2" style="width:6%;">Tổng Xuất</th>
                <th rowspan="2" style="width:6%;">Tồn cuối tháng</th>
                <th rowspan="2" style="width:8%;">Ghi chú</th>
            </tr>
            <tr>${destRow}</tr>`;
    }

    function setStatus(msg, isErr) {
        const el = document.getElementById("thxl-appsheet-status");
        if (!el) return;
        el.textContent = msg;
        el.style.color = isErr ? "#b91c1c" : "#166534";
    }

    function renderTable() {
        const tbody = document.getElementById("thxl-tbody");
        if (!tbody) return;

        const selectedDateRange = getSelectedDateRangeFromDom();
        const tonDkAgg = buildTonKhoDkAggregateByTenNl(cacheTonKhoDkRows, selectedDateRange);

        const useNxlc = Array.isArray(cacheNxlcCtRows) && !lastNxlcFetchError;
        const nhapSlAgg = useNxlc ? buildNxlcSlAggregateFromNxlc(cacheNxlcCtRows, nxlcRowIsPhieuNhap) : null;
        const xuatByDest = useNxlc ? buildXuatSlByDestinationFromNxlc(cacheNxlcCtRows) : new Map();
        const ghiMap = useNxlc ? buildGhiChuMapFromNxlc(cacheNxlcCtRows) : new Map();

        const destinations = collectXuatDestinationsSorted(xuatByDest);
        renderThead(destinations);
        const destCols = destinations.length ? destinations : ["(Không có phiếu xuất theo bộ lọc)"];

        let viewFuelRows = cacheFuelRows || [];
        const nhomNlSel = document.getElementById("thxl-filter-nhom-nl")?.value?.trim() ?? "";
        if (nhomNlSel) {
            viewFuelRows = viewFuelRows.filter((r) => pickDsNhomNhienLieuPhu(r).trim() === nhomNlSel);
        }
        const tenNlSub = document.getElementById("thxl-filter-ten-nl")?.value?.trim().toLowerCase() ?? "";
        if (tenNlSub) {
            viewFuelRows = viewFuelRows.filter((r) => pickDsTenNhienLieu(r).toLowerCase().includes(tenNlSub));
        }

        const sorted = [...viewFuelRows].sort((a, b) => {
            const ma = pickDsMaThietBi(a).localeCompare(pickDsMaThietBi(b), "vi");
            if (ma !== 0) return ma;
            return pickDsTenNhienLieu(a).localeCompare(pickDsTenNhienLieu(b), "vi");
        });

        let html = "";
        let idx = 0;
        const totals = {
            tonDau: 0,
            nhap: 0,
            xuatByDest: new Map(),
            tongXuat: 0,
            tonCuoi: 0
        };

        for (const r of sorted) {
            idx += 1;
            const tenNl = pickDsTenNhienLieu(r);
            const nk = normalizeKeyPart(tenNl);
            const dvt = pickDsDvt(r);

            let tonDauSl = null;
            const hit = tonDkAgg.get(nk);
            if (hit && hit.sl != null && !isNaN(hit.sl)) tonDauSl = hit.sl;

            let nhapSl = 0;
            if (useNxlc && nhapSlAgg) {
                const v = getAggValueForDsFromNxlcMap(nhapSlAgg, tenNl);
                nhapSl = v != null && !isNaN(v) ? v : 0;
            }

            const xuatCells = [];
            let tongXuat = 0;
            for (let di = 0; di < destCols.length; di++) {
                const dest = destCols[di];
                const realDest = destinations.length ? destinations[di] : null;
                let v = 0;
                if (realDest && useNxlc) {
                    const raw = getXuatAtDest(xuatByDest, realDest, tenNl);
                    v = raw != null && !isNaN(raw) ? raw : 0;
                }
                tongXuat += v;
                totals.xuatByDest.set(dest, (totals.xuatByDest.get(dest) || 0) + v);
                xuatCells.push(`<td class="num">${escapeHtml(displaySlWithDvtNl(v, dvt))}</td>`);
            }

            const tonCuoiSl = computeClosingValue(tonDauSl, nhapSl, tongXuat);

            if (tonDauSl != null && !isNaN(tonDauSl)) totals.tonDau += tonDauSl;
            totals.nhap += nhapSl;
            totals.tongXuat += tongXuat;
            if (tonCuoiSl != null && !isNaN(tonCuoiSl)) totals.tonCuoi += tonCuoiSl;

            const gc = mergeGhiChuSetsForDs(ghiMap, tenNl);

            html += `<tr class="thxl-data-row">
                <td>${idx}</td>
                <td>${escapeHtml(tenNl)}</td>
                <td>${escapeHtml(dvt)}</td>
                <td class="num">${escapeHtml(displaySlWithDvtNl(tonDauSl, dvt))}</td>
                <td class="num">${escapeHtml(displaySlWithDvtNl(nhapSl, dvt))}</td>
                ${xuatCells.join("")}
                <td class="num">${escapeHtml(displaySlWithDvtNl(tongXuat, dvt))}</td>
                <td class="num">${escapeHtml(displaySlWithDvtNl(tonCuoiSl, dvt))}</td>
                <td style="text-align:left;font-size:12px;">${escapeHtml(gc)}</td>
            </tr>`;
        }

        if (!sorted.length) {
            const colspan = 5 + destCols.length + 3;
            html = `<tr><td colspan="${colspan}" style="padding:12px;">Không có dòng «Danh sách tài sản» nào có «Tên nhóm» = «${escapeHtml(
                NHOM_NHIEN_LIEU_LABEL
            )}».</td></tr>`;
        } else {
            const tf = (n) => (n != null && !isNaN(n) ? formatNum(n) : "");
            const xuatTotCells = destCols.map((dest) => {
                const v = totals.xuatByDest.get(dest) || 0;
                return `<td class="num">${tf(v)}</td>`;
            });
            html += `<tr class="thxl-total-row" style="font-weight:800;background:#e5e7eb;">
                <td></td>
                <td style="text-align:center;">TỔNG CỘNG</td>
                <td></td>
                <td class="num">${tf(totals.tonDau)}</td>
                <td class="num">${tf(totals.nhap)}</td>
                ${xuatTotCells.join("")}
                <td class="num">${tf(totals.tongXuat)}</td>
                <td class="num">${tf(totals.tonCuoi)}</td>
                <td></td>
            </tr>`;
        }

        tbody.innerHTML = html;
    }

    async function loadFromAppSheet(forceRefresh) {
        const btn = document.getElementById("thxl-btn-load");
        const refBtn = document.getElementById("thxl-btn-refresh");
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
            populateThxlNxlcFilters(cacheNxlcCtRows || []);
            syncThxlBanners();
            cacheFuelRows = (cacheDsRows || []).filter(rowIsNhomNhienLieu);
            populateThxlNhomNlFilter(cacheFuelRows);
            renderTable();
            const n = cacheDsRows?.length ?? 0;
            const m = cacheFuelRows.length;
            const t = cacheTonKhoDkRows?.length ?? 0;
            const x = cacheNxlcCtRows?.length ?? 0;
            let statusMsg = `Đã tải ${n} dòng «${TABLE_DS_TAI_SAN}», ${t} dòng «${TABLE_TON_KHO_DK}», ${x} dòng «${TABLE_NXLC_CT}» — ${m} dòng nhóm «${NHOM_NHIEN_LIEU_LABEL}».`;
            if (nxlcLoadWarn) {
                statusMsg += ` — lỗi NXLC: ${nxlcLoadWarn}`;
            }
            setStatus(statusMsg, !!nxlcLoadWarn);
        } catch (e) {
            console.error(e);
            setStatus(`Lỗi: ${e.message}`, true);
            cacheFuelRows = [];
            populateThxlNhomNlFilter([]);
            renderTable();
        } finally {
            if (btn) btn.disabled = false;
            if (refBtn) refBtn.disabled = false;
        }
    }

    function init() {
        const loadBtn = document.getElementById("thxl-btn-load");
        const refBtn = document.getElementById("thxl-btn-refresh");
        if (loadBtn) loadBtn.addEventListener("click", () => loadFromAppSheet(false));
        if (refBtn) refBtn.addEventListener("click", () => loadFromAppSheet(true));

        const search = document.getElementById("thxl-search");
        if (search) {
            search.addEventListener("input", () => {
                const q = search.value.trim().toLowerCase();
                const tbody = document.getElementById("thxl-tbody");
                if (!tbody) return;
                for (const tr of tbody.querySelectorAll("tr.thxl-data-row")) {
                    tr.style.display = !q || tr.textContent.toLowerCase().includes(q) ? "" : "none";
                }
            });
        }

        const rerenderTableView = () => {
            syncThxlBanners();
            renderTable();
        };
        for (const id of ["thxl-filter-from-date", "thxl-filter-to-date"]) {
            const el = document.getElementById(id);
            if (!el) continue;
            el.addEventListener("change", rerenderTableView);
            el.addEventListener("input", rerenderTableView);
        }
        for (const id of ["thxl-filter-ncc", "thxl-filter-cong-trinh", "thxl-filter-ten-kho", "thxl-filter-nhom-nl"]) {
            document.getElementById(id)?.addEventListener("change", () => {
                if (id === "thxl-filter-cong-trinh") syncThxlCongTrinhBanner();
                rerenderTableView();
            });
        }
        document.getElementById("thxl-filter-ten-nl")?.addEventListener("input", rerenderTableView);

        wireNativeDatePickerButton("thxl-filter-from-date", "thxl-filter-from-date-picker", "thxl-filter-from-date-cal");
        wireNativeDatePickerButton("thxl-filter-to-date", "thxl-filter-to-date-picker", "thxl-filter-to-date-cal");

        for (const id of ["thxl-filter-from-date", "thxl-filter-to-date"]) {
            const el = document.getElementById(id);
            if (!el) continue;
            el.addEventListener("input", syncThxlBanners);
            el.addEventListener("keyup", syncThxlBanners);
            el.addEventListener("change", syncThxlBanners);
            el.addEventListener("blur", syncThxlBanners);
            el.addEventListener("paste", () => setTimeout(syncThxlBanners, 0));
        }
        syncThxlBanners();

        loadFromAppSheet(false);
    }

    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
    else init();
})();
