/**
 * Kết nối AppSheet cho «Bảng kê chi tiết nhập xuất nhiên liệu».
 * Dữ liệu chi tiết: bảng «Nhập xuất luân chuyển CT» — gom theo «Tên XMTB»,
 * tổng «Số lượng NL» theo «Tên nhiên liệu» → 5 cột báo cáo.
 * «Công trình» = cột «Tên nơi quản lý» trên «Nhật trình máy CT» (khớp Tên tài sản/XMTB ↔ Tên thiết bị UI qua DS).
 * «Tên nơi quản lý» (cột LM) = «Tên LM» trên «Danh sách tài sản».
 * «Tồn đầu kỳ» (dòng tổng): «Tồn kho ĐK» — cột «Sl tồn ĐK», lọc «Ngày nhập số dư ĐK» theo khoảng Từ ngày..Đến ngày.
 */
(function () {
    "use strict";

    const APPSHEET_CONFIG = {
        appId: "3be5baea-960f-4d3f-b388-d13364cc4f22",
        accessKey: "V2-GaoRd-ItaM1-r44oH-c6Smd-uOe7V-cmVoK-IJINF-5XLQa"
    };

    const TABLE_NT_CT = "Nhật trình máy CT";
    const TABLE_NT_NL = "Nhật trình NL";
    const TABLE_DS_TAI_SAN = "Danh sách tài sản";
    const TABLE_NXLC_CT = "Nhập xuất luân chuyển CT";
    const TABLE_TON_KHO_DK = "Tồn kho ĐK";

    const NT_FILTER_DATE_COLUMNS = ["Ngày làm việc", "Ngay lam viec", "Ngày", "Ngay"];
    const CT_FILTER_DATE_COLUMNS = ["Ngày làm việc", "Ngay lam viec", "Ngày", "Ngay"];

    const FILTER_DROPDOWN_SOURCE_KEYS = {
        /** Công trình / công trường (NXLC CT — lọc & sổ); không gộp «Tên nơi quản lý». */
        project: [
            "Tên công trình",
            "TenCongTrinh",
            "Ten cong trinh",
            "Công trường",
            "CongTruong",
            "Cong truong",
            "Công trình",
            "Cong trinh",
            "CongTrinh",
            "Tên CT",
            "Ten CT",
            "Dự án",
            "Du an"
        ],
        /** Nơi thi công — cha; «CÔNG TRÌNH» sổ theo giá trị đã chọn. */
        noiThiCong: [
            "Nơi thi công",
            "Noi thi cong",
            "NoiThiCong",
            "Khu vực thi công",
            "Khu vuc thi cong",
            "Địa điểm thi công",
            "Dia diem thi cong"
        ],
        group: ["Nhóm xe máy thiết bị", "NhomXeMayThietBi", "Tên nhóm", "Ten nhom"],
        driver: ["Tên lái máy", "TenLaiMay", "Tên lái xe", "Ten lai xe", "Ten lai may"],
        origin: ["Nguồn gốc", "Nguon goc", "Nguồn gốc xe", "Nguon goc xe", "Origin"]
    };

    const REPORT_FIELD_ALIASES = {
        TenThietBi: ["Tên tài sản", "Ten tai san", "TenThietBi", "Tên thiết bị", "Ten thiet bi"],
        /** Cột «Tên LM» trên «Danh sách tài sản» — hiển thị ở cột Người phụ trách. */
        TenLm: ["Tên LM", "Ten LM", "TenLM", "TênLM"],
        NguoiPhuTrach: [
            "Người phụ trách",
            "Nguoi phu trach",
            "NguoiPhuTrach",
            "Phụ trách",
            "Phu trach",
            "CB phụ trách",
            "CB phu trach",
            "Cán bộ phụ trách",
            "Can bo phu trach"
        ],
        CongTruong: [
            "CongTruong",
            "Công trường",
            "Cong truong",
            "Tên công trình",
            "Công trình",
            "Cong trinh",
            "CongTrinh"
        ],
        TonDauKy: ["TonDauKy", "Tồn đầu kỳ", "Ton dau ky"],
        LuongDauNhap: ["LuongDauNhap", "Lượng dầu xuất", "Luong dau xuat", "Lượng dầu nhập", "Luong dau nhap"],
        TonCuoiKy: ["TonCuoiKy", "Tồn cuối kỳ", "Ton cuoi ky"],
        LuongDauTieuHao: ["LuongDauTieuHao", "Lượng dầu tiêu hao", "Luong dau tieu hao"],
        TongMo: ["TongMo", "Tổng mỡ", "Tong mo"],
        TongNhot: ["TongNhot", "Tổng nhớt", "Tong nhot"],
        DauThuyLuc: ["DauThuyLuc", "Dầu thủy lực", "Dau thuy luc"],
        DauCau: ["DauCau", "Dầu cầu", "Dau cau"],
        DinhMucTrongBinh: ["Định mức trong bình", "Dinh muc trong binh", "DinhMucTrongBinh"]
    };

    const REPORT_COLUMNS = [
        "TenThietBi",
        "NguoiPhuTrach",
        "CongTruong",
        "TonDauKy",
        "LuongDauNhap",
        "TonCuoiKy",
        "LuongDauTieuHao",
        "TongMo",
        "TongNhot",
        "DauThuyLuc",
        "DauCau",
        "DinhMucTrongBinh"
    ];

    const NHOM_TONG_SUM_LITER_COLS = ["LuongDauTieuHao", "TongMo", "TongNhot", "DauThuyLuc", "DauCau"];
    const NXLC_AUX_FUEL_COLUMNS = ["TongMo", "TongNhot", "DauThuyLuc", "DauCau"];
    const NL_FUEL_COLUMN_KEYS = ["TongMo", "TongNhot", "DauThuyLuc", "DauCau"];

    const TEN_NHOM_PRIORITY_KEYS = [
        "Tên nhóm",
        "Ten nhom",
        "TenNhom",
        "Tên nhóm xe",
        "Ten nhom xe",
        "TenNhomXe",
        "Nhóm xe",
        "Nhom xe",
        "Group",
        "group"
    ];

    const NHOM_UNKNOWN_LABEL = "Không phân nhóm";

    /** Cache một lần toàn bộ dòng «Nhập xuất luân chuyển CT» (Làm mới = tải lại). */
    let cacheNxlcCtRows = null;
    /** Cache «Danh sách tài sản» — tra «Tên LM» theo Tên tài sản ↔ Tên thiết bị (fallback). */
    let cacheDsTaiSanRows = null;
    /** Cache «Nhật trình máy CT» — cột «Tên nơi quản lý», khớp «Tên tài sản» với «Tên thiết bị» báo cáo. */
    let cacheNtCtRows = null;
    /** Cache «Nhật trình NL» — «Tồn trước khi nhập» / Ngày cho dòng tổng «Tồn đầu kỳ». */
    let cacheNtNlRows = null;
    /** Cache «Tồn kho ĐK» — nguồn dòng tổng «Tồn đầu kỳ». */
    let cacheTonKhoDkRows = null;
    const NL_PREVIEW_ROW_CAP = 500;

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

    function normalizeKeyPart(s) {
        return String(s ?? "")
            .trim()
            .normalize("NFD")
            .replace(/\u0300-\u036f/g, "")
            .toLowerCase()
            .replace(/[^a-z0-9]/g, "");
    }

    function parentIdsMatch(ref, parentId) {
        const a0 = String(ref == null ? "" : ref).trim();
        const b0 = String(parentId == null ? "" : parentId).trim();
        if (!a0 || !b0) return false;
        if (a0.toLowerCase() === b0.toLowerCase()) return true;
        const na = normalizeKeyPart(a0);
        const nb = normalizeKeyPart(b0);
        if (na === nb) return true;
        const MIN = 8;
        if (na.length >= MIN && nb.length >= MIN && (na.includes(nb) || nb.includes(na))) return true;
        if (nb.length >= MIN && na.length < nb.length && nb.includes(na)) return true;
        if (na.length >= MIN && nb.length < na.length && na.includes(nb)) return true;
        const lastA = a0.lastIndexOf("_");
        const lastB = b0.lastIndexOf("_");
        const tailA = lastA >= 0 ? normalizeKeyPart(a0.slice(lastA + 1)) : na;
        const tailB = lastB >= 0 ? normalizeKeyPart(b0.slice(lastB + 1)) : nb;
        if (tailA && tailB) {
            if (tailA === tailB) return true;
            if (tailA.length >= MIN && tailB.length >= MIN && (tailA.includes(tailB) || tailB.includes(tailA)))
                return true;
        }
        const ax = a0.toLowerCase();
        const bx = b0.toLowerCase();
        if (ax.includes(bx) || bx.includes(ax)) {
            if (Math.min(ax.length, bx.length) >= 6) return true;
        }
        return false;
    }

    function assetKeyFromName(name) {
        return normalizeKeyPart(String(name ?? "").trim());
    }

    function pickField(raw, columnKey) {
        const keys = REPORT_FIELD_ALIASES[columnKey] || [columnKey];
        for (const k of keys) {
            const v = raw[k];
            if (v === undefined || v === null) continue;
            if (typeof v === "string" && v.trim() === "") continue;
            return v;
        }
        return "";
    }

    function parseNumericForLiters(raw) {
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

    function formatLitersDisplay(n) {
        if (n == null || isNaN(n)) return "";
        return new Intl.NumberFormat("vi-VN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
    }

    function escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    function isLikelyDateTimeColumnKey(k) {
        const s = String(k ?? "");
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

    function cellValueToDdMmYyyyStr(v) {
        if (v == null || v === "") return "";
        if (v instanceof Date && !isNaN(v.getTime())) return formatDateVn(v);
        const s = cellDisplayString(v).trim();
        if (!s) return "";
        const d = parseDateFlexible(s);
        return d && !isNaN(d.getTime()) ? formatDateVn(d) : "";
    }

    function orderPreviewColumnKeys(keys) {
        const all = keys.filter(Boolean);
        const pri = all.filter(isLikelyDateTimeColumnKey).sort((a, b) => a.localeCompare(b, "vi"));
        const rest = all.filter((k) => !isLikelyDateTimeColumnKey(k)).sort((a, b) => a.localeCompare(b, "vi"));
        return [...pri, ...rest];
    }

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

    function previewCellValueForModal(v, columnKey) {
        if (v == null) return "";
        const col = columnKey != null ? String(columnKey) : "";
        const rawDisp = cellDisplayString(v).trim();
        const shouldFormatDate =
            (col && isLikelyDateTimeColumnKey(col)) || (rawDisp && looksLikeRawDateString(rawDisp));
        if (shouldFormatDate) {
            const dd = cellValueToDdMmYyyyStr(v);
            if (dd) return dd.length > 800 ? `${dd.slice(0, 800)}...` : dd;
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
        return s.length > 800 ? `${s.slice(0, 800)}...` : s;
    }

    function buildRawPreviewTableHtml(rows) {
        const r = rows ?? [];
        if (!r.length) return '<p class="nl-prev-note">Không có dòng.</p>';
        const total = r.length;
        const show = r.slice(0, NL_PREVIEW_ROW_CAP);
        const keys = collectKeysFromRows(r);
        if (!keys.length) return `<p class="nl-prev-note">${total} dòng - không có khóa cột.</p>`;
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
            (total > NL_PREVIEW_ROW_CAP ? `Hiển thị tối đa ${NL_PREVIEW_ROW_CAP} / ${total} dòng. ` : `${total} dòng. `) +
            "Cột ngày/giờ xếp trước; giá trị ngày hiển thị dd/mm/yyyy. Kéo ngang nếu bảng rộng.";
        return `<p class="nl-prev-note">${escapeHtml(note)}</p><div class="nl-prev-scroll"><table class="nl-prev-table"><thead>${thead}</thead><tbody>${tbody}</tbody></table></div>`;
    }

    function formatPreviewTabCount(filteredCount, totalCount) {
        if (filteredCount === totalCount) return String(totalCount);
        return `${filteredCount}/${totalCount}`;
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
        // Tránh JS tự đoán chuỗi số kiểu mm/dd.
        // Chỉ fallback native khi chuỗi có chữ (tháng text/timezone text).
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

    function dateKeyYmd(d) {
        if (!d || isNaN(d.getTime())) return "";
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const dd = String(d.getDate()).padStart(2, "0");
        return `${yyyy}-${mm}-${dd}`;
    }

    function appSheetLooksLikeDateString(raw) {
        const s = String(raw ?? "").trim();
        if (!s) return false;
        if (/^\d{4}[/-]\d{1,2}[/-]\d{1,2}(?:[T\s,].*)?$/i.test(s)) return true;
        if (/^\d{1,2}[/-]\d{1,2}[/-]\d{4}(?:[T\s,].*)?$/i.test(s)) return true;
        return false;
    }

    function normalizeAppSheetDateCell(value, keyHint) {
        const keyLooksDate = /ng[aà]y|date/i.test(String(keyHint ?? ""));
        const normalizeString = (s) => {
            const text = String(s ?? "").trim();
            if (!text) return s;
            if (!keyLooksDate && !appSheetLooksLikeDateString(text)) return s;
            if (keyLooksDate) {
                const mUs = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:[\s,].*)?$/);
                if (mUs) {
                    const mm = parseInt(mUs[1], 10);
                    const dd = parseInt(mUs[2], 10);
                    const yy = parseInt(mUs[3], 10);
                    const dUs = buildStrictDate(yy, mm, dd);
                    if (dUs && !isNaN(dUs.getTime())) return formatDateVn(dUs);
                }
            }
            const d = parseDateFlexible(text);
            if (!d || isNaN(d.getTime())) return s;
            return formatDateVn(d);
        };
        if (typeof value === "string") return normalizeString(value);
        if (value instanceof Date) return formatDateVn(value);
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

    /** Chỉ loại trùng theo «Row ID» AppSheet (không dùng ID phiếu / _RowNumber — trùng trên nhiều dòng chi tiết). */
    function dedupeNxlcCtRowsByRowIdentity(rows) {
        if (!rows?.length) return rows;
        const seen = new Set();
        const out = [];
        for (const r of rows) {
            const id = cellString(r["Row ID"] ?? r.RowId ?? "").trim();
            if (id) {
                const key = id.replace(/\s+/g, " ").toLowerCase();
                if (seen.has(key)) continue;
                seen.add(key);
            }
            out.push(r);
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

    async function fetchAuxiliaryTablesParallel() {
        const [rCt, rNl, rDs, rNxlcCt] = await Promise.allSettled([
            fetchAppSheetTable(TABLE_NT_CT),
            fetchAppSheetTable(TABLE_NT_NL),
            fetchAppSheetTable(TABLE_DS_TAI_SAN),
            fetchAppSheetTable(TABLE_NXLC_CT)
        ]);
        const u = (res, label) => {
            if (res.status === "fulfilled") return { rows: res.value, error: null };
            console.warn(label, res.reason);
            return { rows: [], error: res.reason };
        };
        return {
            ctRows: u(rCt, "NT CT").rows,
            ctLoadError: u(rCt, "").error,
            nlRows: u(rNl, "NT NL").rows,
            nlLoadError: u(rNl, "").error,
            dsTaiSanRows: u(rDs, "DS").rows,
            dsTaiSanLoadError: u(rDs, "").error,
            nxlcCtRows: dedupeNxlcCtRowsByRowIdentity(u(rNxlcCt, "NXLC CT").rows),
            nxlcCtLoadError: u(rNxlcCt, "").error
        };
    }

    function getNtParentMatchIds(record) {
        const ids = [];
        const seen = new Set();
        const add = (v) => {
            const t = cellString(v).trim();
            if (t === "" || seen.has(t)) return;
            seen.add(t);
            ids.push(t);
        };
        if (!record || typeof record !== "object") return ids;
        for (const k of ["ID", "Row ID", "RowId", "_RowNumber", "ID phát hiện", "Id phat hien", "KEY", "Key"]) {
            add(record[k]);
        }
        return ids;
    }

    function getCtParentRefCandidates(row) {
        if (!row || typeof row !== "object") return [];
        const orderedKeys = [
            "ID cha",
            "ID Cha",
            "Id cha",
            "ID phiếu",
            "ID phieu",
            "Ref",
            "REF",
            "Nhật trình máy",
            "Nhat trinh may"
        ];
        const found = [];
        const seen = new Set();
        for (const k of orderedKeys) {
            const t = cellString(row[k]).trim();
            if (t === "" || seen.has(t)) continue;
            seen.add(t);
            found.push(t);
        }
        for (const [k, v] of Object.entries(row)) {
            const t = cellString(v).trim();
            if (t === "") continue;
            if (/phiếu|phieu|ref|nhật trình máy|id cha|related/i.test(k)) {
                if (!seen.has(t)) {
                    seen.add(t);
                    found.push(t);
                }
            }
        }
        return found;
    }

    function getNlParentRefCandidates(row) {
        if (!row || typeof row !== "object") return [];
        const orderedKeys = [
            "ID phiếu",
            "ID phieu",
            "Ref",
            "REF",
            "Nhật trình máy",
            "ID cha",
            "ID Cha"
        ];
        const found = [];
        const seen = new Set();
        for (const k of orderedKeys) {
            const t = cellString(row[k]).trim();
            if (t === "" || seen.has(t)) continue;
            seen.add(t);
            found.push(t);
        }
        for (const [k, v] of Object.entries(row)) {
            const t = cellString(v).trim();
            if (t === "") continue;
            if (/phiếu|phieu|ref|nhật trình máy|id cha|related/i.test(k)) {
                if (!seen.has(t)) {
                    seen.add(t);
                    found.push(t);
                }
            }
        }
        return found;
    }

    function parseRecordNgayLamViec(raw) {
        if (!raw || typeof raw !== "object") return null;
        let v = "";
        for (const k of NT_FILTER_DATE_COLUMNS) {
            const t = cellString(raw[k]).trim();
            if (t) {
                v = t;
                break;
            }
        }
        if (!v) return null;
        const d = parseDateFlexible(v);
        return d && !isNaN(d.getTime()) ? d : null;
    }

    function pickCtTaiSanTen(row) {
        if (!row || typeof row !== "object") return "";
        return String(
            row["Tên tài sản"] ??
                row["Ten tai san"] ??
                row["TenTaiSan"] ??
                row["Tên thiết bị"] ??
                row["Ten thiet bi"] ??
                ""
        ).trim();
    }

    /** «Nhật trình máy CT» có thể có «Tên XMTB» song song «Tên tài sản». */
    function pickCtTenXmtbTen(row) {
        if (!row || typeof row !== "object") return "";
        const keys = ["Tên XMTB", "Ten XMTB", "TenXMTB", "XMTB", "Tên xe máy thiết bị", "Ten xe may thiet bi"];
        for (const k of keys) {
            const v = cellDisplayString(row[k] ?? "").trim();
            if (v) return v;
        }
        return "";
    }

    function taiSanNamesLooselyEqual(a, b) {
        const s1 = String(a ?? "").trim();
        const s2 = String(b ?? "").trim();
        if (!s1 || !s2) return false;
        const n1 = normalizeKeyPart(s1);
        const n2 = normalizeKeyPart(s2);
        if (n1 && n2 && n1 === n2) return true;
        const c1 = s1.toLowerCase();
        const c2 = s2.toLowerCase();
        if (c1 === c2) return true;
        if (n1.length >= 5 && n2.length >= 5 && (n1.includes(n2) || n2.includes(n1))) return true;
        return false;
    }

    function normTextForXeBonMatch(s) {
        return String(s ?? "")
            .trim()
            .normalize("NFKC")
            .normalize("NFD")
            .replace(/\u0300-\u036f/g, "")
            .toLowerCase();
    }

    function nhomTenLooksLikeXeBon(s) {
        const raw = String(s ?? "").trim();
        if (!raw) return false;
        const t = normTextForXeBonMatch(raw);
        if (/\bxe\s*bon\b/.test(t) || t.includes("xe bon")) return true;
        if (/\bxebon\b/.test(t) || t.includes("xe-bon")) return true;
        if (/\bnhom\s*xe\s*bon\b/.test(t) || t.includes("nhom xe bon")) return true;
        if (/\bbon\b/.test(t) && /\bxe\b/.test(t)) return true;
        const o = raw.normalize("NFKC");
        if (/xe\s*bồn|Xe\s*bồn|Bồn\s*xe/i.test(o)) return true;
        return false;
    }

    function collectTenNhomCandidatesFromRow(raw) {
        if (!raw || typeof raw !== "object") return [];
        const seen = new Set();
        const out = [];
        const add = (v) => {
            const t = cellString(v).trim();
            if (!t || seen.has(t)) return;
            seen.add(t);
            out.push(t);
        };
        const keyOrder = [...(FILTER_DROPDOWN_SOURCE_KEYS.group || []), ...TEN_NHOM_PRIORITY_KEYS];
        for (const k of keyOrder) {
            if (raw[k] == null || raw[k] === "") continue;
            add(raw[k]);
        }
        for (const [k, v] of Object.entries(raw)) {
            const kn = String(k).trim();
            if (/^t[eê]n\s*nh[oô]m(\s|$)/i.test(kn) || /^ten\s*nhom(\s|$)/i.test(kn.replace(/\s+/g, " "))) add(v);
        }
        return out;
    }

    function pickPreferredTenNhomFromCandidates(candidates) {
        if (!candidates?.length) return "";
        for (const c of candidates) {
            if (nhomTenLooksLikeXeBon(c)) return c;
        }
        return candidates[0];
    }

    function pickCtTenNhom(row) {
        return pickPreferredTenNhomFromCandidates(collectTenNhomCandidatesFromRow(row));
    }

    function pickDanhSachTaiSanTen(r) {
        return cellDisplayString(
            r["Tên tài sản"] ?? r["Ten tai san"] ?? r["TenTaiSan"] ?? r["Tên thiết bị"] ?? r["Ten thiet bi"] ?? ""
        ).trim();
    }

    /** Trên DS: cùng dòng tài sản có thể có «Tên XMTB» khác «Tên tài sản» — cần để khớp với báo cáo gom theo XMTB. */
    function pickDanhSachTenXmtb(r) {
        if (!r || typeof r !== "object") return "";
        const keys = ["Tên XMTB", "Ten XMTB", "TenXMTB", "XMTB", "Tên xe máy thiết bị", "Ten xe may thiet bi"];
        for (const k of keys) {
            const v = cellDisplayString(r[k] ?? "").trim();
            if (v) return v;
        }
        return "";
    }

    /**
     * Mọi nhãn có thể cùng một máy (tên hiển thị báo cáo, Tên tài sản DS, Tên XMTB DS) — dùng khi tra NT CT / DS.
     */
    function expandAssetLabelsFromDanhSachTaiSan(tenThietBi, dsRows) {
        const tx = String(tenThietBi ?? "").trim();
        const out = new Set();
        if (!tx) return out;
        out.add(tx);
        if (!dsRows?.length) return out;
        for (const r of dsRows) {
            const tten = pickDanhSachTaiSanTen(r);
            const xmtb = pickDanhSachTenXmtb(r);
            const parts = [tten, xmtb].filter(Boolean);
            if (!parts.some((p) => taiSanNamesLooselyEqual(tx, p))) continue;
            if (tten) out.add(tten);
            if (xmtb) out.add(xmtb);
        }
        return out;
    }

    function pickDanhSachNguonGoc(r) {
        const keys = ["Nguồn gốc", "Nguon goc", "Nguồn gốc xe", "Nguon goc xe", "Origin"];
        for (const k of keys) {
            const v = cellDisplayString(r[k] ?? "").trim();
            if (v) return v;
        }
        return "";
    }

    function pickDanhSachNguoiPhuTrach(r) {
        if (!r || typeof r !== "object") return "";
        for (const k of REPORT_FIELD_ALIASES.NguoiPhuTrach || []) {
            const v = cellDisplayString(r[k] ?? "").trim();
            if (v) return v;
        }
        for (const [k, v] of Object.entries(r)) {
            const kn = String(k)
                .normalize("NFD")
                .replace(/\u0300-\u036f/g, "")
                .toLowerCase();
            if (!/phu\s*trach|cb\s*phu/.test(kn)) continue;
            const t = cellDisplayString(v).trim();
            if (t) return t;
        }
        return "";
    }

    /** «Tên nơi quản lý» trên «Nhật trình máy CT» — khớp tài sản qua pickCtTaiSanTen. */
    function pickNtCtTenNoiQuanLy(r) {
        if (!r || typeof r !== "object") return "";
        const keys = [
            "Tên nơi quản lý",
            "Ten noi quan ly",
            "TenNoiQuanLy",
            "Tên nơi quản lí",
            "Noi quan ly"
        ];
        for (const k of keys) {
            const v = cellDisplayString(r[k] ?? "").trim();
            if (v) return v;
        }
        for (const [k, v] of Object.entries(r)) {
            const kn = String(k)
                .normalize("NFD")
                .replace(/\u0300-\u036f/g, "")
                .toLowerCase();
            if (!/noi\s*quan\s*ly|quan\s*ly\s*tai\s*san/.test(kn)) continue;
            const t = cellDisplayString(v).trim();
            if (t) return t;
        }
        return "";
    }

    function buildTenNoiQuanLyByTaiSanFromNtCt(ntRows) {
        const map = new Map();
        if (!ntRows?.length) return map;
        for (const r of ntRows) {
            const nq = pickNtCtTenNoiQuanLy(r);
            if (!nq) continue;
            for (const label of [pickCtTaiSanTen(r), pickCtTenXmtbTen(r)]) {
                if (!label) continue;
                const key = assetKeyFromName(label);
                if (key) map.set(key, nq);
            }
        }
        return map;
    }

    /** Khớp «Tên thiết bị» báo cáo (thường là XMTB) với «Nhật trình máy CT».Tên tài sản → «Tên nơi quản lý». */
    function resolveTenNoiQuanLyForTenThietBi(tenThietBi, noiQuanLyByKey, ntRows, dsRows) {
        const tx = String(tenThietBi ?? "").trim();
        if (!tx) return "";
        const labels = expandAssetLabelsFromDanhSachTaiSan(tx, dsRows);
        for (const lab of labels) {
            const k = assetKeyFromName(lab);
            if (k && noiQuanLyByKey?.has(k)) return noiQuanLyByKey.get(k);
        }
        if (!ntRows?.length) return "";
        for (const r of ntRows) {
            const candidates = [pickCtTaiSanTen(r), pickCtTenXmtbTen(r)].filter(Boolean);
            if (!candidates.length) continue;
            let match = false;
            outer: for (const cand of candidates) {
                for (const lab of labels) {
                    if (taiSanNamesLooselyEqual(lab, cand)) {
                        match = true;
                        break outer;
                    }
                }
            }
            if (!match) continue;
            const nq = pickNtCtTenNoiQuanLy(r);
            if (nq) return nq;
        }
        return "";
    }

    /** «Tên LM» trên DS — khớp «Tên tài sản» (DS) với «Tên thiết bị» báo cáo (Tên XMTB). */
    function pickDanhSachTenLm(r) {
        if (!r || typeof r !== "object") return "";
        for (const k of REPORT_FIELD_ALIASES.TenLm || []) {
            const v = cellDisplayString(r[k] ?? "").trim();
            if (v) return v;
        }
        for (const [k, v] of Object.entries(r)) {
            const kn = String(k).trim();
            if (!/^t[eê]n\s*lm$/i.test(kn) && !/^ten\s*lm$/i.test(kn)) continue;
            const t = cellDisplayString(v).trim();
            if (t) return t;
        }
        return "";
    }

    function buildTenLmByTaiSanFromDanhSachTaiSan(dsRows) {
        const map = new Map();
        if (!dsRows?.length) return map;
        for (const r of dsRows) {
            const lm = pickDanhSachTenLm(r);
            if (!lm) continue;
            const name = pickDanhSachTaiSanTen(r);
            if (name) {
                const key = assetKeyFromName(name);
                if (key) map.set(key, lm);
            }
            const xm = pickDanhSachTenXmtb(r);
            if (xm) {
                const k2 = assetKeyFromName(xm);
                if (k2) map.set(k2, lm);
            }
        }
        return map;
    }

    /** Khớp «Tên XMTB» với «Danh sách tài sản».Tên tài sản → «Tên LM». */
    function resolveTenLmForTenXmtb(tenXmtb, tenLmByKey, dsRows) {
        const tx = String(tenXmtb ?? "").trim();
        if (!tx) return "";
        const labels = expandAssetLabelsFromDanhSachTaiSan(tx, dsRows);
        for (const lab of labels) {
            const k = assetKeyFromName(lab);
            if (k && tenLmByKey?.has(k)) return tenLmByKey.get(k);
        }
        if (!dsRows?.length) return "";
        for (const r of dsRows) {
            const name = pickDanhSachTaiSanTen(r);
            if (!name) continue;
            let match = false;
            for (const lab of labels) {
                if (taiSanNamesLooselyEqual(lab, name)) {
                    match = true;
                    break;
                }
            }
            if (!match) continue;
            const lm = pickDanhSachTenLm(r);
            if (lm) return lm;
        }
        return "";
    }

    function buildNguonGocByTaiSanFromDanhSachTaiSan(dsRows) {
        const map = new Map();
        if (!dsRows?.length) return map;
        for (const r of dsRows) {
            const name = pickDanhSachTaiSanTen(r);
            if (!name) continue;
            const ng = pickDanhSachNguonGoc(r);
            if (!ng) continue;
            const key = assetKeyFromName(name);
            if (!key) continue;
            map.set(key, ng);
        }
        return map;
    }

    function buildTenLoaiByTaiSanFromDanhSachTaiSan(dsRows) {
        const map = new Map();
        if (!dsRows?.length) return map;
        return map;
    }

    function applyNgayLamViecFilterByCt(mainRows, ctRows, filters) {
        const from = parseDateInputYmd(filters?.fromDate);
        const to = parseDateInputYmd(filters?.toDate);
        if (!from && !to) return mainRows || [];
        if (!mainRows?.length) return [];

        const fallbackByMainDate = (rows) =>
            (rows || []).filter((raw) => {
                const d = parseRecordNgayLamViec(raw);
                if (!d || isNaN(d.getTime())) return false;
                if (from && d < from) return false;
                if (to && d > to) return false;
                return true;
            });

        if (!ctRows?.length) return fallbackByMainDate(mainRows);

        const mainAssetByIndex = mainRows.map((raw) => String(pickField(raw, "TenThietBi") ?? "").trim());
        const mainByAssetKey = new Map();
        const mainIdsByIndex = mainRows.map((raw, index) => ({ index, ids: getNtParentMatchIds(raw) }));
        for (let i = 0; i < mainRows.length; i += 1) {
            const k = assetKeyFromName(mainAssetByIndex[i]);
            if (!k) continue;
            if (!mainByAssetKey.has(k)) mainByAssetKey.set(k, []);
            mainByAssetKey.get(k).push(i);
        }

        const matchedMainIndexes = new Set();
        let ctRowsInRange = 0;

        for (const ct of ctRows) {
            ctRowsInRange += 1;

            const ctAsset = pickCtTaiSanTen(ct);
            const ctAssetKey = assetKeyFromName(ctAsset);
            if (ctAssetKey && mainByAssetKey.has(ctAssetKey)) {
                for (const idx of mainByAssetKey.get(ctAssetKey)) matchedMainIndexes.add(idx);
            } else if (ctAsset) {
                for (let i = 0; i < mainAssetByIndex.length; i += 1) {
                    if (taiSanNamesLooselyEqual(ctAsset, mainAssetByIndex[i])) matchedMainIndexes.add(i);
                }
            }

            const refs = getCtParentRefCandidates(ct);
            if (!refs.length) continue;
            for (const rref of refs) {
                if (!rref) continue;
                for (const x of mainIdsByIndex) {
                    if (!x.ids?.length) continue;
                    for (const pid of x.ids) {
                        if (parentIdsMatch(rref, pid)) {
                            matchedMainIndexes.add(x.index);
                            break;
                        }
                    }
                }
            }
        }

        if (ctRowsInRange === 0) return fallbackByMainDate(mainRows);
        if (!matchedMainIndexes.size) return fallbackByMainDate(mainRows);
        return mainRows.filter((_, idx) => matchedMainIndexes.has(idx));
    }

    function parseCtNgayLamViec(raw) {
        if (!raw || typeof raw !== "object") return null;
        for (const k of CT_FILTER_DATE_COLUMNS) {
            const t = cellDisplayString(raw[k]).trim();
            if (!t) continue;
            const d = parseDateFlexible(t);
            if (d && !isNaN(d.getTime())) return d;
        }
        for (const [k, v] of Object.entries(raw)) {
            if (!/ng[aà]y|date/i.test(k)) continue;
            const d = parseDateFlexible(cellDisplayString(v).trim());
            if (d && !isNaN(d.getTime())) return d;
        }
        return null;
    }

    function filterCtRowsByNgayLamViec(ctRows, filters) {
        const from = parseDateInputYmd(filters?.fromDate);
        const to = parseDateInputYmd(filters?.toDate);
        if (!from && !to) return ctRows || [];
        if (!ctRows?.length) return [];
        return ctRows.filter((ct) => {
            const d = parseCtNgayLamViec(ct);
            if (!d || isNaN(d.getTime())) return false;
            if (from && d < from) return false;
            if (to && d > to) return false;
            return true;
        });
    }

    function pickAssetNameForGroupSummary(raw) {
        return String(
            raw?.["Tên tài sản"] ??
                raw?.["Ten tai san"] ??
                raw?.["Tên thiết bị"] ??
                raw?.["Ten thiet bi"] ??
                ""
        ).trim();
    }

    function rawCellMatchesNeedle(raw, filterKey, needle) {
        const n = needle.trim().toLowerCase();
        if (!n) return true;
        const cols = FILTER_DROPDOWN_SOURCE_KEYS[filterKey] || [];
        for (const col of cols) {
            const v = cellString(raw[col] ?? "").trim();
            if (!v) continue;
            if (v.toLowerCase() === n) return true;
            if (v.toLowerCase().includes(n)) return true;
        }
        return false;
    }

    function pickFirstDisplayFromColumnKeys(raw, keys) {
        if (!raw || !keys?.length) return "";
        for (const col of keys) {
            const v = cellDisplayString(raw[col] ?? "").trim();
            if (v) return v;
        }
        return "";
    }

    /** Danh sách nhãn «Công trình» (cột «Tên nơi quản lý» trên «Nhật trình máy CT») cho checkbox. */
    function buildNlCongTrinhFilterOptionsFromNtCt(ntRows) {
        const set = new Set();
        for (const r of ntRows || []) {
            const v = pickNtCtTenNoiQuanLy(r);
            if (v) set.add(String(v).trim());
        }
        return [...set].sort((a, b) => a.localeCompare(b, "vi"));
    }

    function populateNlCongTrinhCheckboxList(ntRows) {
        const panel = document.getElementById("nl-filter-cong-trinh-panel");
        if (!panel) return;
        const opts = buildNlCongTrinhFilterOptionsFromNtCt(ntRows);
        const prevChecked = new Set(
            [...panel.querySelectorAll('input[type="checkbox"][name="nl-cong-trinh"]:checked')].map((el) => el.value)
        );
        if (!opts.length) {
            panel.innerHTML = '<span class="text-slate-400">Không có nhãn công trình trên «Nhật trình máy CT»</span>';
            syncNlCongTrinhTitleBanner();
            return;
        }
        const html = opts
            .map((label, i) => {
                const id = `nl-ct-cb-${i}`;
                const checked = prevChecked.has(label) ? " checked" : "";
                return `<label class="flex items-start gap-2 py-0.5 cursor-pointer select-none" for="${id}">
<input type="checkbox" name="nl-cong-trinh" id="${id}" value="${escapeHtml(label)}"${checked}>
<span class="font-normal leading-tight">${escapeHtml(label)}</span>
</label>`;
            })
            .join("");
        panel.innerHTML = html;
        syncNlCongTrinhTitleBanner();
    }

    function getNlSelectedCongTrinhValues() {
        const panel = document.getElementById("nl-filter-cong-trinh-panel");
        if (!panel) return [];
        return [...panel.querySelectorAll('input[type="checkbox"][name="nl-cong-trinh"]:checked')]
            .map((el) => String(el.value ?? "").trim())
            .filter(Boolean);
    }

    /** Dòng «CT …» dưới tiêu đề: theo công trình đã tick; không tick = mặc định cả hai công trường. */
    const NL_DEFAULT_BANNER_SITE_LINE = "CT HỐ SÓI - SÔNG ĐUỐNG";

    function syncNlCongTrinhTitleBanner() {
        const el = document.getElementById("nl-banner-cong-trinh-line");
        if (!el) return;
        const selected = getNlSelectedCongTrinhValues();
        if (!selected.length) {
            el.textContent = NL_DEFAULT_BANNER_SITE_LINE;
            return;
        }
        if (selected.length === 1) {
            el.textContent = selected[0];
            return;
        }
        el.textContent = selected.join(" • ");
    }

    function filterNlRowsBySelectedCongTrinh(rows, selectedLabels) {
        if (!selectedLabels?.length || !rows?.length) return rows || [];
        return rows.filter((row) => {
            const ct = String(row.CongTrinh ?? "").trim();
            if (!ct) return false;
            return selectedLabels.some((needle) => needle === ct || taiSanNamesLooselyEqual(needle, ct));
        });
    }

    function rawRowMatchesOriginFilter(raw, needle, nguonGocByTaiSan) {
        const n = needle.trim().toLowerCase();
        if (!n) return true;
        const asset = pickAssetNameForGroupSummary(raw) || String(pickField(raw, "TenThietBi") ?? "").trim();
        const key = assetKeyFromName(asset);
        const fromDs = key && nguonGocByTaiSan?.get ? nguonGocByTaiSan.get(key) : "";
        if (fromDs) {
            const g = String(fromDs).trim().toLowerCase();
            if (g === n) return true;
            if (g.includes(n) || n.includes(g)) return true;
        }
        return rawCellMatchesNeedle(raw, "origin", needle);
    }

    function rawRowMatchesGroupFilter(raw, needle, tenNhomByTaiSan) {
        const n = needle.trim().toLowerCase();
        if (!n) return true;
        const asset = pickAssetNameForGroupSummary(raw);
        const key = assetKeyFromName(asset);
        const fromCt = key && tenNhomByTaiSan?.get ? tenNhomByTaiSan.get(key) : "";
        if (fromCt) {
            const g = String(fromCt).trim().toLowerCase();
            if (g === n) return true;
            if (g.includes(n)) return true;
        }
        return rawCellMatchesNeedle(raw, "group", needle);
    }

    function applyRawTextFilters(rows, filters, tenNhomByTaiSan, nguonGocByTaiSan) {
        if (!rows?.length) return rows || [];
        return rows.filter((r) => {
            for (const fk of ["project", "driver"]) {
                const needle = (filters[fk] ?? "").trim();
                if (!needle) continue;
                if (!rawCellMatchesNeedle(r, fk, needle)) return false;
            }
            const oNeedle = (filters.origin ?? "").trim();
            if (oNeedle && !rawRowMatchesOriginFilter(r, oNeedle, nguonGocByTaiSan)) return false;
            const gNeedle = (filters.group ?? "").trim();
            if (gNeedle && !rawRowMatchesGroupFilter(r, gNeedle, tenNhomByTaiSan)) return false;
            return true;
        });
    }

    function pickModeTenNhomFromVoteMap(voteMap) {
        if (!voteMap?.size) return "";
        let best = "";
        let bestN = -1;
        for (const [name, n] of voteMap) {
            if (n > bestN || (n === bestN && name.localeCompare(best, "vi") < 0)) {
                best = name;
                bestN = n;
            }
        }
        return best;
    }

    function buildTenNhomByTaiSanFromNtCt(mainRows, ctRows) {
        const byKey = new Map();
        if (!ctRows?.length || !mainRows?.length) return byKey;

        const mainAssetKeys = new Set();
        for (const raw of mainRows) {
            if (raw?.groupRow || raw?.subGroupRow) continue;
            const asset = String(pickField(raw, "TenThietBi") ?? "").trim();
            if (!asset) continue;
            const key = assetKeyFromName(asset);
            if (!key) continue;
            mainAssetKeys.add(key);
        }

        function addVote(tally, assetKey, tenNhom) {
            const g = String(tenNhom ?? "").trim();
            if (!g || !assetKey) return;
            if (!tally.has(assetKey)) tally.set(assetKey, new Map());
            const m = tally.get(assetKey);
            m.set(g, (m.get(g) ?? 0) + 1);
        }

        const tallyTaiSan = new Map();

        for (const r of ctRows) {
            const tenNhom = pickCtTenNhom(r);
            if (!tenNhom) continue;

            const ctTen = pickCtTaiSanTen(r);
            const k = assetKeyFromName(ctTen);
            if (!k || !mainAssetKeys.has(k)) continue;
            addVote(tallyTaiSan, k, tenNhom);
        }

        for (const assetKey of mainAssetKeys) {
            const fromTaiSan = pickModeTenNhomFromVoteMap(tallyTaiSan.get(assetKey));
            if (fromTaiSan) byKey.set(assetKey, fromTaiSan);
        }
        return byKey;
    }

    function parseNlQuantityCell(row) {
        const keys = [
            "SL nhập",
            "SL nhap",
            "Số lượng nhập",
            "So luong nhap",
            "Số lượng",
            "Khối lượng",
            "Lượng",
            "SL",
            "Sl"
        ];
        for (const k of keys) {
            const q = parseNumericForLiters(row[k]);
            if (q != null) return q;
        }
        for (const [k, v] of Object.entries(row)) {
            if (!/sl\s*nh[aạ]p|lượng|luong|^\s*sl\s*$/i.test(k)) continue;
            const q = parseNumericForLiters(v);
            if (q != null) return q;
        }
        return null;
    }

    function parseNlTonTruocCell(row) {
        const keys = ["Tồn trước khi nhập", "Ton truoc khi nhap", "Tồn trước", "Ton truoc"];
        for (const k of keys) {
            const q = parseNumericForLiters(row[k]);
            if (q != null) return q;
        }
        if (row && typeof row === "object") {
            for (const [k, v] of Object.entries(row)) {
                const kn = String(k)
                    .normalize("NFD")
                    .replace(/\u0300-\u036f/g, "")
                    .toLowerCase();
                if (!/ton\s*truoc|ton\s*dau/.test(kn) || !/nh[aạ]p|khi/.test(kn)) continue;
                const q = parseNumericForLiters(v);
                if (q != null) return q;
            }
        }
        return null;
    }

    function parseNlTonSauCell(row) {
        const keys = ["Tồn sau khi nhập", "Ton sau khi nhap", "Tồn sau", "Ton sau"];
        for (const k of keys) {
            const q = parseNumericForLiters(row[k]);
            if (q != null) return q;
        }
        return null;
    }

    function parseNlNgayCell(row) {
        const keys = ["Ngày", "Ngay", "Ngày nhập", "Ngày làm việc"];
        for (const k of keys) {
            const t = cellString(row[k]).trim();
            if (!t) continue;
            const d = parseDateFlexible(t);
            if (d && !isNaN(d.getTime())) return d;
        }
        for (const [k, v] of Object.entries(row)) {
            if (!/ng[aà]y|date/i.test(k)) continue;
            const d = parseDateFlexible(cellString(v).trim());
            if (d && !isNaN(d.getTime())) return d;
        }
        return null;
    }

    function classifyNlTenToFuelColumn(tenRaw) {
        const s = normalizeTenNhienLieuLabel(tenRaw);
        if (!s) return null;
        const noAccent = s.normalize("NFD").replace(/\u0300-\u036f/g, "").toLowerCase();
        const compact = noAccent.replace(/[^a-z0-9]/g, "");
        if (s.includes("dầu thủy lực") || noAccent.includes("dau thuy luc") || compact.includes("thuyluc")) return "DauThuyLuc";
        if (s.includes("dầu cầu") || noAccent.includes("dau cau") || compact.includes("daucau") || /nhot\s*cau|nhotcau/.test(noAccent))
            return "DauCau";
        // «Dầu Diezel» (AppSheet) — trước mỡ/nhớt
        if (
            compact.includes("daudiezel") ||
            compact.includes("daudiesel") ||
            /dau\s*diezel|dau\s*diesel/i.test(noAccent) ||
            noAccent.includes("diezel") ||
            noAccent.includes("diesel")
        ) return "LuongDauNhap";
        if (/\bdo\b/i.test(noAccent)) return "LuongDauNhap";
        if (compact.includes("gasoil") || compact.includes("hsd")) return "LuongDauNhap";
        if (/^mỡ$/i.test(s) || compact === "mo") return "TongMo";
        if (/^nhớt$/i.test(s) || compact === "nhot") return "TongNhot";
        if (s.includes("mỡ") || noAccent.includes(" mo") || noAccent.startsWith("mo")) return "TongMo";
        if (s.includes("nhớt") || noAccent.includes("nhot")) return "TongNhot";
        if ((noAccent.includes("dau") || s.includes("dầu")) && /\bdo\b/i.test(noAccent)) return "LuongDauNhap";
        return null;
    }

    function getNlTaiSanCell(r) {
        return r["Tên tài sản"] ?? r["Ten tai san"] ?? r["TenTaiSan"] ?? r["Tên thiết bị"] ?? "";
    }

    /** Tên hiển thị trên «Nhật trình NL» để khớp báo cáo (tài sản + XMTB nếu có). */
    function getNlJournalAssetLabels(r) {
        const out = [];
        const a = cellString(getNlTaiSanCell(r)).trim();
        if (a) out.push(a);
        const xm = cellDisplayString(
            r["Tên XMTB"] ?? r["Ten XMTB"] ?? r["TenXMTB"] ?? r["XMTB"] ?? r["Tên xe máy thiết bị"] ?? ""
        ).trim();
        if (xm) out.push(xm);
        return out;
    }

    function findReportAssetKeyForNlJournalRow(nlRow, reportRows, dsRows) {
        if (!nlRow || !reportRows?.length) return "";
        const nlNames = getNlJournalAssetLabels(nlRow);
        if (!nlNames.length) return "";
        for (const raw of reportRows) {
            const ten = String(raw.TenThietBi ?? "").trim();
            if (!ten) continue;
            const labels = expandAssetLabelsFromDanhSachTaiSan(ten, dsRows);
            for (const nlN of nlNames) {
                for (const lab of labels) {
                    if (taiSanNamesLooselyEqual(lab, nlN)) return assetKeyFromName(ten);
                }
            }
        }
        return "";
    }

    function parseTonKhoDkNgayCellCandidates(raw) {
        function uniqueValidDates(list) {
            const out = [];
            const seen = new Set();
            for (const d of list) {
                if (!d || isNaN(d.getTime())) continue;
                const k = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
                if (seen.has(k)) continue;
                seen.add(k);
                out.push(d);
            }
            return out;
        }

        function parseAppSheetLikeDateCandidates(v) {
            if (v == null || v === "") return [];
            if (v instanceof Date && !isNaN(v.getTime())) return [v];
            const s = cellDisplayString(v).trim();
            if (!s) return [];
            let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
            if (m) {
                const d = buildStrictDate(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));
                return d ? [d] : [];
            }
            m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})(?:[\s,].*)?$/);
            if (m) {
                const a = parseInt(m[1], 10);
                const b = parseInt(m[2], 10);
                const y = parseInt(m[3], 10);
                const cands = [];
                if (a > 12 && b <= 12) cands.push(buildStrictDate(y, b, a)); // d/m
                else if (a <= 12 && b > 12) cands.push(buildStrictDate(y, a, b)); // m/d
                else {
                    // Mơ hồ: giữ cả 2 cách hiểu để khớp range linh hoạt.
                    cands.push(buildStrictDate(y, b, a)); // d/m
                    cands.push(buildStrictDate(y, a, b)); // m/d
                }
                return uniqueValidDates(cands);
            }
            const d = parseDateFlexible(s);
            return d && !isNaN(d.getTime()) ? [d] : [];
        }

        return uniqueValidDates(parseAppSheetLikeDateCandidates(raw));
    }

    function parseTonKhoDkNgayCell(row) {
        const keys = [
            "Ngày nhập số dư ĐK",
            "Ngay nhap so du DK",
            "Ngày nhập số dư DK",
            "Ngay nhap so du ĐK",
            "Ngày nhập SD ĐK",
            "Ngay nhap SD DK",
            "Ngày",
            "Ngay"
        ];
        for (const k of keys) {
            const cands = parseTonKhoDkNgayCellCandidates(row?.[k]);
            if (cands.length) return cands[0];
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (!/ng[aà]y|date/i.test(String(k))) continue;
            const cands = parseTonKhoDkNgayCellCandidates(v);
            if (cands.length) return cands[0];
        }
        return null;
    }

    function tonKhoDkRowDateInRange(row, dateRange) {
        if (!dateRange) return true;
        const keys = [
            "Ngày nhập số dư ĐK",
            "Ngay nhap so du DK",
            "Ngày nhập số dư DK",
            "Ngay nhap so du ĐK",
            "Ngày nhập SD ĐK",
            "Ngay nhap SD DK",
            "Ngày",
            "Ngay"
        ];
        for (const k of keys) {
            const cands = parseTonKhoDkNgayCellCandidates(row?.[k]);
            if (cands.some((d) => rowDateInRange(d, dateRange))) return true;
            if (cands.length) return false;
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (!/ng[aà]y|date/i.test(String(k))) continue;
            const cands = parseTonKhoDkNgayCellCandidates(v);
            if (cands.some((d) => rowDateInRange(d, dateRange))) return true;
            if (cands.length) return false;
        }
        return false;
    }

    function parseTonKhoDkTenNlCell(row) {
        const keys = [
            "Tên NL",
            "Ten NL",
            "Tên nhiên liệu",
            "Ten nhien lieu",
            "Mã NL",
            "Ma NL",
            "Mã nhiên liệu",
            "Ma nhien lieu",
            "Loại nhiên liệu",
            "Loai nhien lieu"
        ];
        for (const k of keys) {
            const t = cellDisplayString(row?.[k]).trim();
            if (t) return t;
        }
        for (const [k, v] of Object.entries(row || {})) {
            const kn = String(k)
                .normalize("NFD")
                .replace(/\u0300-\u036f/g, "")
                .toLowerCase();
            if (!/ten\s*nl|ten\s*nhien\s*lieu|loai\s*nhien\s*lieu|ma\s*nl|ma\s*nhien\s*lieu/.test(kn)) continue;
            const t = cellDisplayString(v).trim();
            if (t) return t;
        }
        return "";
    }

    function parseTonKhoDkSlTonCell(row) {
        const keys = ["Sl tồn ĐK", "SL tồn ĐK", "Sl tồn DK", "SL tồn DK", "SL ton DK", "Sl ton DK"];
        for (const k of keys) {
            const q = parseNumericForLiters(row?.[k]);
            if (q != null && !isNaN(q)) return q;
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (!/sl\s*t[oô]n\s*[đd]k|sl\s*ton\s*dk/i.test(String(k))) continue;
            const q = parseNumericForLiters(v);
            if (q != null && !isNaN(q)) return q;
        }
        return null;
    }

    function classifyTonKhoDkTenNlToReportColumn(tenRaw) {
        const s = normalizeTenNhienLieuLabel(tenRaw);
        if (!s) return null;
        const n = s
            .normalize("NFD")
            .replace(/\u0300-\u036f/g, "")
            .toLowerCase()
            .replace(/\s+/g, " ")
            .trim();
        const compact = n.replace(/[^a-z0-9]/g, "");

        // Khớp đúng theo rule user chốt cho «Tồn kho ĐK».
        if (n === "dau diezel" || compact === "daudiezel" || compact === "daudiesel") return "LuongDauTieuHao";
        if (n === "mo" || compact === "mo") return "TongMo";
        if (n === "nhot" || compact === "nhot") return "TongNhot";
        if (n === "dau thuy luc" || compact === "dauthuyluc") return "DauThuyLuc";
        if (n === "dau cau" || compact === "daucau") return "DauCau";
        return null;
    }

    /**
     * Tổng «Tồn đầu kỳ» theo 5 cột nhiên liệu từ «Tồn kho ĐK».
     * Điều kiện: «Ngày nhập số dư ĐK» nằm trong khoảng Từ ngày..Đến ngày.
     */
    function buildTonDauKyTotalsFromTonKhoDk(tonKhoDkRows, filters) {
        const keys = ["LuongDauTieuHao", "TongMo", "TongNhot", "DauThuyLuc", "DauCau"];
        const empty = () => ({
            LuongDauTieuHao: 0,
            TongMo: 0,
            TongNhot: 0,
            DauThuyLuc: 0,
            DauCau: 0,
            hasAny: false
        });
        const dateRange = buildDateRangeFromFilters(filters);
        if (!dateRange || !tonKhoDkRows?.length) return empty();

        const totals = empty();
        for (const r of tonKhoDkRows) {
            if (!tonKhoDkRowDateInRange(r, dateRange)) continue;

            const ton = parseTonKhoDkSlTonCell(r);
            if (ton == null || isNaN(ton)) continue;

            const tenNl = parseTonKhoDkTenNlCell(r);
            const col = classifyTonKhoDkTenNlToReportColumn(tenNl);
            if (!col) continue;
            if (!keys.includes(col)) continue;
            totals[col] += ton;
            totals.hasAny = true;
        }

        return totals;
    }

    /**
     * Tổng «Nhập trong tháng» theo 5 cột nhiên liệu từ «Nhập xuất luân chuyển CT».
     * Điều kiện: Loại phiếu = Nhập, month(Ngày) = month(Từ ngày), và khớp thiết bị đang có trên báo cáo.
     */
    function buildNhapTrongThangTotalsFromNxlcCt(reportRows, nxlcCtRows, filters) {
        const empty = () => ({
            LuongDauTieuHao: 0,
            TongMo: 0,
            TongNhot: 0,
            DauThuyLuc: 0,
            DauCau: 0,
            hasAny: false
        });
        const fromD = parseDateInputYmd(filters?.fromDate);
        if (!fromD || !nxlcCtRows?.length || !reportRows?.length) return empty();

        const totals = empty();
        const mainAssetKeys = new Set();
        const mainAssetPairs = [];
        for (const raw of reportRows) {
            const ten = String(raw?.TenThietBi ?? "").trim();
            if (!ten) continue;
            const k = assetKeyFromName(ten);
            if (!k) continue;
            mainAssetKeys.add(k);
            mainAssetPairs.push({ key: k, name: ten });
        }
        if (!mainAssetKeys.size) return totals;

        function resolveMainAssetKeyByName(candidateName) {
            const t = String(candidateName ?? "").trim();
            if (!t) return "";
            const exact = assetKeyFromName(t);
            if (exact && mainAssetKeys.has(exact)) return exact;
            for (const p of mainAssetPairs) {
                if (taiSanNamesLooselyEqual(t, p.name)) return p.key;
            }
            return "";
        }

        for (const r of nxlcCtRows) {
            const d = parseNxlcCtNgayCell(r);
            if (!d || isNaN(d.getTime())) continue;
            if (d.getFullYear() !== fromD.getFullYear() || d.getMonth() !== fromD.getMonth()) continue;

            const loaiPhieu = getLoaiPhieuOnRow(r);
            if (!loaiPhieuIsNhap(loaiPhieu)) continue;

            const assetKey = resolveNxlcCtRowToMainAssetKey(r, resolveMainAssetKeyByName, mainAssetKeys);
            if (!assetKey) continue;

            const tenNl = getNxlcTenNhienLieuCell(r);
            const col = classifyNxlcTenNhienLieuToReportFiveColumns(tenNl);
            if (!col) continue;

            const q = parseNxlcSlCell(r);
            if (q == null || isNaN(q)) continue;

            totals[col] += q;
            totals.hasAny = true;
        }
        return totals;
    }

    function getNxlcTaiSanCell(r) {
        if (!r || typeof r !== "object") return "";
        const direct = cellDisplayString(
            r["Tên tài sản"] ?? r["Ten tai san"] ?? r["TenTaiSan"] ?? r["Tên thiết bị"] ?? r["Ten thiet bi"] ?? ""
        ).trim();
        if (direct) return direct;
        for (const [k, v] of Object.entries(r)) {
            const kn = String(k)
                .trim()
                .normalize("NFD")
                .replace(/\u0300-\u036f/g, "")
                .toLowerCase();
            if (!/t[eê]n\s*t[aà]i\s*s[aả]n|t[eê]n\s*thi[eê]t\s*b[iị]|ten\s*tai\s*san|ten\s*thiet\s*bi/.test(kn)) continue;
            const t = cellDisplayString(v).trim();
            if (t) return t;
        }
        return "";
    }

    function getNxlcTenXmtbCell(r) {
        const direct = cellDisplayString(
            r["Tên XMTB"] ?? r["Ten XMTB"] ?? r["TenXMTB"] ?? r["XMTB"] ?? r["Tên xe máy thiết bị"] ?? ""
        ).trim();
        if (direct) return direct;
        for (const [k, v] of Object.entries(r || {})) {
            const kn = String(k)
                .trim()
                .normalize("NFD")
                .replace(/\u0300-\u036f/g, "")
                .toLowerCase();
            if (!/xmtb|xe\s*may\s*thiet\s*bi/.test(kn)) continue;
            const t = cellDisplayString(v).trim();
            if (t) return t;
        }
        return "";
    }

    /**
     * Tên hiển thị / khóa gom thiết bị trên «Nhập xuất luân chuyển CT»:
     * ưu tiên «Tên XMTB» (đúng quy tắc gom NL), nếu trống thì «Tên tài sản» / «Tên thiết bị»
     * để cột «Công trình» khớp cùng dòng với «Tên thiết bị» khi AppSheet không điền XMTB.
     */
    function getNxlcDeviceLabelForRow(r) {
        const xmtb = getNxlcTenXmtbCell(r);
        if (xmtb) return xmtb;
        const ts = getNxlcTaiSanCell(r);
        if (ts) return ts;
        return "";
    }

    /** Khớp khóa báo cáo: thử «Tên XMTB» rồi «Tên tài sản» (xe ben: vd. «Xe ben 01» thường ở Tên tài sản). */
    function resolveNxlcCtRowToMainAssetKey(r, resolveMainAssetKeyByName, mainAssetKeys) {
        const labels = [];
        const x = String(getNxlcTenXmtbCell(r) ?? "").trim();
        const ts = String(getNxlcTaiSanCell(r) ?? "").trim();
        if (x) labels.push(x);
        if (ts) labels.push(ts);
        const seen = new Set();
        for (const lab of labels) {
            const norm = lab.replace(/\s+/g, " ").trim();
            if (!norm) continue;
            const lk = norm.toLowerCase();
            if (seen.has(lk)) continue;
            seen.add(lk);
            const key = resolveMainAssetKeyByName(norm);
            if (key && (!mainAssetKeys.size || mainAssetKeys.has(key))) return key;
        }
        return "";
    }

    function parseNxlcSlCell(row) {
        const keys = [
            "Số lượng NL",
            "So luong NL",
            "Số lượng nhiên liệu",
            "So luong nhien lieu",
            "Số lượng",
            "So luong",
            "SL",
            "Sl",
            "Lượng",
            "Luong"
        ];
        for (const k of keys) {
            const q = parseNumericForLiters(row[k]);
            if (q != null) return q;
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (!/s[oố]\s*l[ưu]ợ?ng.*nl|so\s*luong.*nl/i.test(k)) continue;
            const q = parseNumericForLiters(v);
            if (q != null) return q;
        }
        return null;
    }

    function getNxlcTenNhienLieuCell(r) {
        return cellDisplayString(
            r["Tên nhiên liệu"] ?? r["Ten nhien lieu"] ?? r["Tên NL"] ?? r["Loại nhiên liệu"] ?? ""
        ).trim();
    }

    /** Chuẩn hóa chuỗi «Tên nhiên liệu» AppSheet: gộp khoảng trắng thừa, trim. */
    function normalizeTenNhienLieuLabel(tenRaw) {
        return cellDisplayString(tenRaw)
            .normalize("NFKC")
            .replace(/\s+/g, " ")
            .trim();
    }

    /**
     * «Tên nhiên liệu» (NXLC CT) → 5 cột: Dầu diezel | Mỡ bò | Nhớt động cơ | Dầu thủy lực | Dầu cầu.
     * AppSheet «Mỡ» → cột Mỡ bò; «Nhớt» → cột Nhớt động cơ (cùng các tên mở rộng có chứa Mỡ/Nhớt, trừ nhánh cầu/thủy lực/diezel).
     * Cột chuẩn diezel: «Dầu Diezel» — khớp sau chuẩn hóa (dau diezel / dau diesel) trước nhánh DO chung.
     */
    function classifyNxlcTenNhienLieuToReportFiveColumns(tenRaw) {
        const sOrig = normalizeTenNhienLieuLabel(tenRaw);
        if (!sOrig) return null;
        const s = sOrig;
        const no = s.normalize("NFD").replace(/\u0300-\u036f/g, "").toLowerCase();
        const c = no.replace(/[^a-z0-9]/g, "");

        if (s.includes("thủy lực") || /thuy\s*luc/.test(no) || c.includes("thuyluc")) return "DauThuyLuc";
        if (s.includes("dầu cầu") || /dau\s*cau|nhot\s*cau/.test(no) || c.includes("daucau")) return "DauCau";
        // AppSheet đặt tên ngắn «Mỡ» / «Nhớt» (sau chuẩn hóa khoảng trắng)
        if (/^mỡ$/i.test(s) || c === "mo") return "TongMo";
        if (/^nhớt$/i.test(s) || c === "nhot") return "TongNhot";
        if (s.includes("mỡ")) return "TongMo";
        if (s.includes("nhớt") || /\bnhot\b/.test(no)) return "TongNhot";
        const noSp = no.replace(/\s+/g, " ").trim();
        if (noSp === "dau diezel" || noSp === "dau diesel") return "LuongDauTieuHao";
        if (c.includes("diezel") || c.includes("diesel") || /\bdo\b/.test(no) || c.includes("daudiezel") || c.includes("daudiesel"))
            return "LuongDauTieuHao";
        if ((no.includes("dau") || s.includes("dầu")) && /\bdo\b/.test(no)) return "LuongDauTieuHao";
        return null;
    }

    function classifyNxlcTenToReportColumn(tenRaw) {
        const s = normalizeTenNhienLieuLabel(tenRaw);
        if (!s) return null;
        const noAccent = s.normalize("NFD").replace(/\u0300-\u036f/g, "").toLowerCase();
        const compact = noAccent.replace(/[^a-z0-9]/g, "");
        if (s.includes("dầu thủy lực") || noAccent.includes("dau thuy luc") || compact.includes("thuyluc")) return "DauThuyLuc";
        if (s.includes("dầu cầu") || noAccent.includes("dau cau") || compact.includes("daucau")) return "DauCau";
        if (
            compact.includes("daudiezel") ||
            compact.includes("daudiesel") ||
            /dau\s*diezel|dau\s*diesel/i.test(noAccent) ||
            noAccent.includes("diezel") ||
            noAccent.includes("diesel")
        ) return "LuongDauTieuHao";
        if (/\bdo\b/i.test(noAccent)) return "LuongDauTieuHao";
        if (compact.includes("gasoil") || compact.includes("hsd")) return "LuongDauTieuHao";
        if (s.includes("mỡ") || noAccent.includes(" mo") || noAccent.startsWith("mo")) return "TongMo";
        if (s.includes("nhớt") || noAccent.includes("nhot")) return "TongNhot";
        if ((noAccent.includes("dau") || s.includes("dầu")) && /\bdo\b/i.test(noAccent)) return "LuongDauTieuHao";
        return null;
    }

    function getNxlcCtParentRefCandidates(row) {
        if (!row || typeof row !== "object") return [];
        const orderedKeys = ["ID phiếu", "ID phieu", "Ref", "REF", "Nhập xuất luân chuyển", "ID cha"];
        const found = [];
        const seen = new Set();
        for (const k of orderedKeys) {
            const t = cellString(row[k]).trim();
            if (t === "" || seen.has(t)) continue;
            seen.add(t);
            found.push(t);
        }
        for (const [k, v] of Object.entries(row)) {
            const t = cellString(v).trim();
            if (t === "") continue;
            if (/phiếu|phieu|ref|nhập xuất|id cha|related/i.test(k) && !seen.has(t)) {
                seen.add(t);
                found.push(t);
            }
        }
        return found;
    }

    /** «Loại phiếu» = «Xuất» (AppSheet) hoặc cụm bắt đầu bằng Xuất / Phiếu xuất. */
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

    function getLoaiPhieuOnRow(row) {
        if (!row || typeof row !== "object") return "";
        const explicitKeys = [
            "Loại phiếu",
            "Loai phieu",
            "Loại phiếu NX",
            "Loại giao dịch",
            "Loại GD",
            "Kiểu phiếu",
            "Kieu phieu",
            "Phân loại phiếu",
            "Phan loai phieu"
        ];
        for (const k of explicitKeys) {
            const v = row[k];
            if (v == null || String(v).trim() === "") continue;
            return v;
        }
        for (const k of ["Loại", "Loai"]) {
            const v = row[k];
            if (v == null || String(v).trim() === "") continue;
            const t = String(v).trim();
            if (loaiPhieuIsXuat(t) || loaiPhieuIsNhap(t)) return v;
        }
        for (const k of ["Phiếu", "Phieu"]) {
            const v = row[k];
            if (v == null || String(v).trim() === "") continue;
            return v;
        }
        for (const [k, v] of Object.entries(row)) {
            if (v == null || String(v).trim() === "") continue;
            if (/lo[aạ]i\s*phi[ếe]u|loai\s*phieu/i.test(k)) return v;
        }
        return "";
    }

    function nxlcCtRowIsPhieuXuat(r) {
        const lp = getLoaiPhieuOnRow(r);
        if (loaiPhieuIsXuat(lp)) return true;
        if (loaiPhieuIsNhap(lp)) return false;
        const nx = cellString(r["Ngày"] ?? r["Ngay"] ?? "").trim();
        const nn = cellString(r["Ngày nhập"] ?? r["Ngay nhap"] ?? "").trim();
        return !!(nx && !nn);
    }

    function getNxlcParentMatchIds(record) {
        const ids = [];
        const seen = new Set();
        const add = (v) => {
            const t = cellString(v).trim();
            if (t === "" || seen.has(t)) return;
            seen.add(t);
            ids.push(t);
        };
        if (!record || typeof record !== "object") return ids;
        for (const k of ["ID", "Row ID", "RowId", "_RowNumber", "KEY", "Key"]) add(record[k]);
        return ids;
    }

    function parseNxlcCtNgayCell(row) {
        // NXLC CT chuẩn hiện tại: chỉ dùng cột «Ngày».
        const keys = ["Ngày", "Ngay"];
        for (const k of keys) {
            const t = cellString(row[k]).trim();
            if (!t) continue;
            const d = parseDateFlexible(t);
            if (d && !isNaN(d.getTime())) return d;
        }
        for (const [k, v] of Object.entries(row)) {
            const kn = String(k ?? "").trim().toLowerCase();
            if (!(kn === "ngày" || kn === "ngay" || kn === "date")) continue;
            const d = parseDateFlexible(cellString(v).trim());
            if (d && !isNaN(d.getTime())) return d;
        }
        return null;
    }

    function emptyNlFuelSumsByTaiSan() {
        return {
            TongMo: new Map(),
            TongNhot: new Map(),
            DauThuyLuc: new Map(),
            DauCau: new Map()
        };
    }

    function buildNlFuelSumsByTaiSan(mainRows, nlRows, filters) {
        const sums = emptyNlFuelSumsByTaiSan();
        if (!nlRows?.length || !mainRows?.length) return sums;
        const from = parseDateInputYmd(filters?.fromDate);
        const to = parseDateInputYmd(filters?.toDate);

        const mainAssetKeys = new Set();
        for (const raw of mainRows) {
            if (raw?.groupRow || raw?.subGroupRow) continue;
            const asset = String(pickField(raw, "TenThietBi") ?? "").trim();
            if (!asset) continue;
            const key = assetKeyFromName(asset);
            if (!key) continue;
            mainAssetKeys.add(key);
        }

        function addToMap(targetMap, assetKey, q) {
            targetMap.set(assetKey, (targetMap.get(assetKey) ?? 0) + q);
        }

        for (const r of nlRows) {
            if (from || to) {
                const d = parseNlNgayCell(r);
                if (!d || isNaN(d.getTime())) continue;
                if (from && d < from) continue;
                if (to && d > to) continue;
            }
            const tenNl =
                r["Tên nhiên liệu"] ?? r["Ten nhien lieu"] ?? r["Loại nhiên liệu"] ?? r["Loai nhien lieu"] ?? "";
            const col = classifyNlTenToFuelColumn(tenNl);
            if (!col || col === "LuongDauNhap") continue;
            const q = parseNlQuantityCell(r);
            if (q == null) continue;

            const targetMap = sums[col];
            const assetDirect = cellString(getNlTaiSanCell(r)).trim();
            const k = assetKeyFromName(assetDirect);
            if (!k || !mainAssetKeys.has(k)) continue;
            addToMap(targetMap, k, q);
        }
        return sums;
    }

    /** Tồn đầu/cuối: ưu tiên dòng diezel/DO (cùng ý báo cáo vật tư xe bồn), không thì mọi loại NL. */
    function buildNlTonBoundaryByTaiSan(mainRows, nlRows, filters) {
        const tonDauByTaiSan = new Map();
        const tonCuoiByTaiSan = new Map();
        if (!nlRows?.length || !mainRows?.length) return { tonDauByTaiSan, tonCuoiByTaiSan };

        const fromD = parseDateInputYmd(filters?.fromDate);
        const toD = parseDateInputYmd(filters?.toDate);
        const fromKey = dateKeyYmd(fromD);
        if (!fromKey && !toD) return { tonDauByTaiSan, tonCuoiByTaiSan };

        const tonDauDateDiesel = new Map();
        const tonDauValDiesel = new Map();
        const tonDauDateAny = new Map();
        const tonDauValAny = new Map();
        const tonCuoiDateDiesel = new Map();
        const tonCuoiValDiesel = new Map();
        const tonCuoiDateAny = new Map();
        const tonCuoiValAny = new Map();

        const parents = [];
        for (const raw of mainRows) {
            if (raw?.groupRow || raw?.subGroupRow) continue;
            const ids = getNtParentMatchIds(raw);
            if (!ids.length) continue;
            const asset = String(pickField(raw, "TenThietBi") ?? "").trim();
            if (!asset) continue;
            const key = assetKeyFromName(asset);
            if (!key) continue;
            parents.push({ ids, key });
        }

        function detectAssetKeyOnNlRow(r) {
            const assetDirect = cellString(getNlTaiSanCell(r)).trim();
            if (assetDirect) {
                const k = assetKeyFromName(assetDirect);
                if (k) return k;
            }
            for (const rref of getNlParentRefCandidates(r)) {
                if (!rref) continue;
                for (const { ids, key } of parents) {
                    for (const pid of ids) {
                        if (parentIdsMatch(rref, pid)) return key;
                    }
                }
            }
            return "";
        }

        function nlLineIsDiesel(r) {
            const tenNl =
                r["Tên nhiên liệu"] ??
                r["Ten nhien lieu"] ??
                r["Tên NL"] ??
                r["Loại nhiên liệu"] ??
                r["Loai nhien lieu"] ??
                "";
            return classifyNlTenToFuelColumn(tenNl) === "LuongDauNhap";
        }

        for (const r of nlRows) {
            const assetKey = detectAssetKeyOnNlRow(r);
            if (!assetKey) continue;
            const ngay = parseNlNgayCell(r);
            const dKey = dateKeyYmd(ngay);
            if (!dKey) continue;

            const isDiesel = nlLineIsDiesel(r);

            if (
                fromD &&
                ngay &&
                ngay.getFullYear() === fromD.getFullYear() &&
                ngay.getMonth() === fromD.getMonth()
            ) {
                const tonTruoc = parseNlTonTruocCell(r);
                if (tonTruoc != null) {
                    const pickedAny = tonDauDateAny.get(assetKey);
                    if (!pickedAny || ngay < pickedAny) {
                        tonDauDateAny.set(assetKey, ngay);
                        tonDauValAny.set(assetKey, tonTruoc);
                    }
                    if (isDiesel) {
                        const pickedD = tonDauDateDiesel.get(assetKey);
                        if (!pickedD || ngay < pickedD) {
                            tonDauDateDiesel.set(assetKey, ngay);
                            tonDauValDiesel.set(assetKey, tonTruoc);
                        }
                    }
                }
            }
            if (
                toD &&
                ngay &&
                ngay.getFullYear() === toD.getFullYear() &&
                ngay.getMonth() === toD.getMonth()
            ) {
                const tonSau = parseNlTonSauCell(r);
                if (tonSau != null) {
                    const pickedAny = tonCuoiDateAny.get(assetKey);
                    if (!pickedAny || ngay > pickedAny) {
                        tonCuoiDateAny.set(assetKey, ngay);
                        tonCuoiValAny.set(assetKey, tonSau);
                    }
                    if (isDiesel) {
                        const pickedD = tonCuoiDateDiesel.get(assetKey);
                        if (!pickedD || ngay > pickedD) {
                            tonCuoiDateDiesel.set(assetKey, ngay);
                            tonCuoiValDiesel.set(assetKey, tonSau);
                        }
                    }
                }
            }
        }

        for (const k of new Set([...tonDauValDiesel.keys(), ...tonDauValAny.keys()])) {
            tonDauByTaiSan.set(
                k,
                tonDauValDiesel.has(k) ? tonDauValDiesel.get(k) : tonDauValAny.get(k)
            );
        }
        for (const k of new Set([...tonCuoiValDiesel.keys(), ...tonCuoiValAny.keys()])) {
            tonCuoiByTaiSan.set(
                k,
                tonCuoiValDiesel.has(k) ? tonCuoiValDiesel.get(k) : tonCuoiValAny.get(k)
            );
        }

        return { tonDauByTaiSan, tonCuoiByTaiSan };
    }

    function mergeFuelMapsByKey(a, b) {
        const out = new Map();
        const keys = new Set([...(a?.keys?.() ?? []), ...(b?.keys?.() ?? [])]);
        for (const k of keys) {
            const s = (a?.get(k) ?? 0) + (b?.get(k) ?? 0);
            out.set(k, s);
        }
        return out;
    }

    function buildMergedAuxiliaryFuelMapsFromNlAndNxlc(nlSums, nxlcSums) {
        const empty = () => new Map();
        return {
            TongMo: mergeFuelMapsByKey(empty(), nxlcSums?.TongMo ?? empty()),
            TongNhot: mergeFuelMapsByKey(empty(), nxlcSums?.TongNhot ?? empty()),
            DauThuyLuc: mergeFuelMapsByKey(empty(), nxlcSums?.DauThuyLuc ?? empty()),
            DauCau: mergeFuelMapsByKey(empty(), nxlcSums?.DauCau ?? empty())
        };
    }

    function buildFuelColumnsByTaiSanFromNxlcCt(nxlcCtRows, mainRows, filters) {
        const sums = {
            LuongDauNhap: new Map(),
            LuongDauTieuHao: new Map(),
            TongMo: new Map(),
            TongNhot: new Map(),
            DauThuyLuc: new Map(),
            DauCau: new Map()
        };
        if (!nxlcCtRows?.length) return sums;
        const from = parseDateInputYmd(filters?.fromDate);
        const to = parseDateInputYmd(filters?.toDate);

        const mainAssetKeys = new Set();
        const mainAssetPairs = [];
        if (mainRows?.length) {
            for (const raw of mainRows) {
                if (raw?.groupRow || raw?.subGroupRow) continue;
                const asset = String(pickField(raw, "TenThietBi") ?? "").trim();
                if (!asset) continue;
                const key = assetKeyFromName(asset);
                if (!key) continue;
                mainAssetKeys.add(key);
                mainAssetPairs.push({ key, name: asset });
            }
        }

        function resolveMainAssetKeyByName(candidateName) {
            const t = String(candidateName ?? "").trim();
            if (!t) return "";
            const exact = assetKeyFromName(t);
            if (exact && (!mainAssetKeys.size || mainAssetKeys.has(exact))) return exact;
            for (const p of mainAssetPairs) {
                if (taiSanNamesLooselyEqual(t, p.name)) return p.key;
            }
            return "";
        }

        function rowIsLoaiPhieuXuat(r) {
            return nxlcCtRowIsPhieuXuat(r);
        }

        const luongDauTieuHaoNlTab = new Map();
        for (const r of nxlcCtRows) {
            const d0 = parseNxlcCtNgayCell(r);
            if (from || to) {
                if (!d0 || isNaN(d0.getTime())) continue;
                if (from && d0 < from) continue;
                if (to && d0 > to) continue;
            }
            const kDiezel = resolveNxlcCtRowToMainAssetKey(r, resolveMainAssetKeyByName, mainAssetKeys);
            if (!kDiezel) continue;
            const col5 = classifyNxlcTenNhienLieuToReportFiveColumns(getNxlcTenNhienLieuCell(r));
            if (col5 !== "LuongDauTieuHao") continue;
            const q0 = parseNxlcSlCell(r);
            if (q0 == null) continue;
            luongDauTieuHaoNlTab.set(kDiezel, (luongDauTieuHaoNlTab.get(kDiezel) ?? 0) + q0);
        }
        sums.LuongDauTieuHao = luongDauTieuHaoNlTab;

        for (const r of nxlcCtRows) {
            const d = parseNxlcCtNgayCell(r);
            if (from || to) {
                if (!d || isNaN(d.getTime())) continue;
                if (from && d < from) continue;
                if (to && d > to) continue;
            }
            const tenNlCell = getNxlcTenNhienLieuCell(r);
            const col5 = classifyNxlcTenNhienLieuToReportFiveColumns(tenNlCell);
            if (col5 === "LuongDauTieuHao" && rowIsLoaiPhieuXuat(r)) {
                const keyXuat = resolveNxlcCtRowToMainAssetKey(r, resolveMainAssetKeyByName, mainAssetKeys);
                if (keyXuat) {
                    const qXuat = parseNxlcSlCell(r);
                    if (qXuat != null) {
                        sums.LuongDauNhap.set(keyXuat, (sums.LuongDauNhap.get(keyXuat) ?? 0) + qXuat);
                    }
                }
                continue;
            }
            const reportCol = classifyNxlcTenToReportColumn(tenNlCell);
            if (!reportCol) continue;
            if (reportCol === "LuongDauTieuHao") continue;
            let asset = "";
            if (NXLC_AUX_FUEL_COLUMNS.includes(reportCol)) {
                asset = cellString(getNxlcTenXmtbCell(r)).trim();
            } else {
                asset = cellString(getNxlcTaiSanCell(r)).trim();
            }
            if (!asset) continue;
            const key = resolveMainAssetKeyByName(asset);
            if (!key || (mainAssetKeys.size && !mainAssetKeys.has(key))) continue;
            const q = parseNxlcSlCell(r);
            if (q == null) continue;
            const targetMap = sums[reportCol];
            targetMap.set(key, (targetMap.get(key) ?? 0) + q);
        }
        return sums;
    }

    function applyLuongDauNhapFromNxlcCt(mainRows, normalizedRows, nxlcFuelSums) {
        const mNx = nxlcFuelSums?.LuongDauNhap;
        const mainAssetPairs = [];
        for (const raw of mainRows || []) {
            if (raw?.groupRow || raw?.subGroupRow) continue;
            const asset = String(pickField(raw, "TenThietBi") ?? "").trim();
            const k = assetKeyFromName(asset);
            if (k) mainAssetPairs.push({ key: k, name: asset });
        }
        function findLuongMapKey(tenThietBi, map) {
            if (!map?.size) return "";
            const k0 = assetKeyFromName(tenThietBi);
            if (k0 && map.has(k0)) return k0;
            const t = String(tenThietBi ?? "").trim();
            for (const p of mainAssetPairs) {
                if (!map.has(p.key)) continue;
                if (taiSanNamesLooselyEqual(t, p.name)) return p.key;
            }
            return "";
        }
        return normalizedRows.map((out, i) => {
            if (out.groupRow || out.subGroupRow || out.nhomTongRow) return out;
            const raw = mainRows[i];
            if (!raw) return out;
            if (!mNx?.size) {
                out.LuongDauNhap = "";
                return out;
            }
            const kNx = findLuongMapKey(out.TenThietBi, mNx);
            out.LuongDauNhap = kNx ? formatLitersDisplay(mNx.get(kNx)) : "";
            return out;
        });
    }

    function applyMergedAuxiliaryFuelColumnsFromNlNxlc(mainRows, normalizedRows, mergedAux) {
        if (!mergedAux) return normalizedRows;
        const hasAny = NXLC_AUX_FUEL_COLUMNS.some((c) => mergedAux[c]?.size > 0);
        if (!hasAny) return normalizedRows;
        return normalizedRows.map((out, i) => {
            if (out.groupRow || out.subGroupRow || out.nhomTongRow) return out;
            const raw = mainRows[i];
            if (!raw) return out;
            const key = assetKeyFromName(out.TenThietBi);
            if (!key) return out;
            for (const c of NXLC_AUX_FUEL_COLUMNS) {
                const map = mergedAux[c];
                if (map?.has(key)) out[c] = formatLitersDisplay(map.get(key));
            }
            return out;
        });
    }

    function applyTonBoundaryColumnsFromNl(mainRows, normalizedRows, tonBoundary, dinhMucMaps) {
        const tonDauMap = tonBoundary?.tonDauByTaiSan;
        const tonCuoiMap = tonBoundary?.tonCuoiByTaiSan;

        function getDinhMucTrongBinhForRow(out, key) {
            const fromDs = dinhMucMaps?.trongBinhByTaiSan?.get(key);
            if (fromDs != null && !isNaN(fromDs) && fromDs > 0) return fromDs;
            const parsed = parseNumericForLiters(out.DinhMucTrongBinh);
            if (parsed != null && !isNaN(parsed) && parsed > 0) return parsed;
            return null;
        }

        function nhanDinhMuc(baseVal, dm) {
            if (baseVal == null || isNaN(baseVal)) return null;
            if (dm == null || isNaN(dm) || dm <= 0) return baseVal;
            return baseVal * dm;
        }

        return normalizedRows.map((out, i) => {
            if (out.groupRow || out.subGroupRow || out.nhomTongRow) return out;
            const raw = mainRows[i];
            if (!raw) return out;
            const key = assetKeyFromName(out.TenThietBi);
            if (!key) return out;

            const dm = getDinhMucTrongBinhForRow(out, key);

            if (tonDauMap?.has(key)) {
                const v = nhanDinhMuc(tonDauMap.get(key), dm);
                if (v != null) out.TonDauKy = formatLitersDisplay(v);
            } else {
                const base = parseNumericForLiters(out.TonDauKy);
                if (base != null) {
                    const v = nhanDinhMuc(base, dm);
                    if (v != null) out.TonDauKy = formatLitersDisplay(v);
                }
            }

            if (tonCuoiMap?.has(key)) {
                const v = nhanDinhMuc(tonCuoiMap.get(key), dm);
                if (v != null) out.TonCuoiKy = formatLitersDisplay(v);
            } else {
                const base = parseNumericForLiters(out.TonCuoiKy);
                if (base != null) {
                    const v = nhanDinhMuc(base, dm);
                    if (v != null) out.TonCuoiKy = formatLitersDisplay(v);
                }
            }
            return out;
        });
    }

    function applyFuelColumnsFromNxlcCt(mainRows, normalizedRows, fuelByTaiSan) {
        if (!fuelByTaiSan) return normalizedRows;
        const m = fuelByTaiSan.LuongDauTieuHao;
        if (!m?.size) return normalizedRows;
        return normalizedRows.map((out, i) => {
            if (out.groupRow || out.subGroupRow || out.nhomTongRow) return out;
            const raw = mainRows[i];
            if (!raw) return out;
            const key = assetKeyFromName(out.TenThietBi);
            if (!key || !m.has(key)) return out;
            out.LuongDauTieuHao = formatLitersDisplay(m.get(key));
            return out;
        });
    }

    function applyLuongDauTieuHaoFromTonFormula(mainRows, normalizedRows, nxlcFuelSums) {
        if (!normalizedRows?.length) return normalizedRows;
        const mainAssetPairs = [];
        for (const raw of mainRows || []) {
            if (raw?.groupRow || raw?.subGroupRow) continue;
            const asset = String(pickField(raw, "TenThietBi") ?? "").trim();
            const k = assetKeyFromName(asset);
            if (k) mainAssetPairs.push({ key: k, name: asset });
        }
        function findLuongMapKey(tenThietBi, map) {
            if (!map?.size) return "";
            const k0 = assetKeyFromName(tenThietBi);
            if (k0 && map.has(k0)) return k0;
            const t = String(tenThietBi ?? "").trim();
            for (const p of mainAssetPairs) {
                if (!map.has(p.key)) continue;
                if (taiSanNamesLooselyEqual(t, p.name)) return p.key;
            }
            return "";
        }
        const mNx = nxlcFuelSums?.LuongDauTieuHao;
        return normalizedRows.map((out, i) => {
            if (out.groupRow || out.subGroupRow || out.nhomTongRow) return out;
            const raw = mainRows?.[i];
            if (!raw) return out;

            const kNx = findLuongMapKey(out.TenThietBi, mNx);
            if (kNx) return out;

            const nTonDau = parseNumericForLiters(out.TonDauKy);
            const nNhap = parseNumericForLiters(out.LuongDauNhap);
            const nTonCuoi = parseNumericForLiters(out.TonCuoiKy);
            if (nTonDau == null || nNhap == null || nTonCuoi == null) return out;

            const tieuHao = nTonDau + nNhap - nTonCuoi;
            out.LuongDauTieuHao = formatLitersDisplay(tieuHao);
            return out;
        });
    }

    function normalizeRowsFromAppSheet(rows) {
        return rows.map((raw) => {
            const out = {};
            for (const col of REPORT_COLUMNS) {
                out[col] = pickField(raw, col);
            }
            if (raw.groupRow) out.groupRow = true;
            if (raw.subGroupRow) out.subGroupRow = true;
            return out;
        });
    }

    function parseCellForNhomTongSum(columnKey, rawVal) {
        return parseNumericForLiters(rawVal);
    }

    function formatNhomTongSum(columnKey, n) {
        if (n == null || isNaN(n)) return "";
        return formatLitersDisplay(n);
    }

    function buildNhomTongAggregateRow(assetRows, groupLabel) {
        const out = {
            nhomTongRow: true,
            groupRow: true,
            TenThietBi: groupLabel,
            NguoiPhuTrach: "",
            CongTruong: ""
        };
        for (const col of NHOM_TONG_SUM_LITER_COLS) {
            let s = 0;
            let any = false;
            for (const r of assetRows) {
                const q = parseCellForNhomTongSum(col, r[col]);
                if (q != null && !isNaN(q)) {
                    s += q;
                    any = true;
                }
            }
            out[col] = any ? formatNhomTongSum(col, s) : "";
        }
        return out;
    }

    function injectNhomTongRows(normalizedRows, tenNhomByTaiSan) {
        if (!normalizedRows?.length || !tenNhomByTaiSan?.size) return normalizedRows;
        const dataRows = normalizedRows.filter((r) => !r.groupRow && !r.subGroupRow && !r.nhomTongRow);
        if (!dataRows.length) return normalizedRows;

        let hasMapped = false;
        for (const r of dataRows) {
            const k = assetKeyFromName(r.TenThietBi);
            if (k && tenNhomByTaiSan.has(k)) {
                hasMapped = true;
                break;
            }
        }
        if (!hasMapped) return normalizedRows;

        const byGroup = new Map();
        for (const r of dataRows) {
            const k = assetKeyFromName(r.TenThietBi);
            const g = (k && tenNhomByTaiSan.get(k)) || NHOM_UNKNOWN_LABEL;
            if (!byGroup.has(g)) byGroup.set(g, []);
            byGroup.get(g).push(r);
        }

        const groupNames = [...byGroup.keys()].sort((a, b) => {
            if (a === NHOM_UNKNOWN_LABEL) return 1;
            if (b === NHOM_UNKNOWN_LABEL) return -1;
            return a.localeCompare(b, "vi");
        });

        const out = [];
        for (const gName of groupNames) {
            const list = byGroup.get(gName);
            list.sort((a, b) => String(a.TenThietBi ?? "").localeCompare(String(b.TenThietBi ?? ""), "vi"));
            out.push(buildNhomTongAggregateRow(list, gName));
            out.push(...list);
        }
        return out;
    }

    /**
     * Một dòng báo cáo = một thiết bị (ưu tiên «Tên XMTB», fallback «Tên tài sản»/«Tên thiết bị»);
     * mỗi ô nhiên liệu = tổng Số lượng NL các dòng CT cùng thiết bị và «Tên nhiên liệu» khớp loại.
     */
    function buildAggregatedRowsFromNxlcCt(nxlcCtRows, filters) {
        const agg = new Map();
        const from = parseDateInputYmd(filters?.fromDate);
        const to = parseDateInputYmd(filters?.toDate);

        for (const r of nxlcCtRows || []) {
            const d = parseNxlcCtNgayCell(r);
            if (from || to) {
                if (!d || isNaN(d.getTime())) continue;
                if (from && d < from) continue;
                if (to && d > to) continue;
            }
            const tenThietBi = getNxlcDeviceLabelForRow(r);
            if (!tenThietBi) continue;

            const key = assetKeyFromName(tenThietBi);
            if (!key) continue;

            const tenNl = getNxlcTenNhienLieuCell(r);
            const col = classifyNxlcTenNhienLieuToReportFiveColumns(tenNl);
            if (!col) continue;

            const q = parseNxlcSlCell(r);
            if (q == null || isNaN(q)) continue;

            if (!agg.has(key)) {
                agg.set(key, {
                    TenThietBi: tenThietBi,
                    LuongDauTieuHao: 0,
                    TongMo: 0,
                    TongNhot: 0,
                    DauThuyLuc: 0,
                    DauCau: 0
                });
            }
            const o = agg.get(key);
            o[col] += q;
        }

        const list = [...agg.values()].sort((a, b) =>
            String(a.TenThietBi ?? "").localeCompare(String(b.TenThietBi ?? ""), "vi")
        );
        return list.map((o) => {
            return {
                TenThietBi: o.TenThietBi,
                NguoiPhuTrach: "",
                CongTrinh: "",
                CongTruong: "",
                LuongDauTieuHao: o.LuongDauTieuHao > 0 ? formatLitersDisplay(o.LuongDauTieuHao) : "",
                TongMo: o.TongMo > 0 ? formatLitersDisplay(o.TongMo) : "",
                TongNhot: o.TongNhot > 0 ? formatLitersDisplay(o.TongNhot) : "",
                DauThuyLuc: o.DauThuyLuc > 0 ? formatLitersDisplay(o.DauThuyLuc) : "",
                DauCau: o.DauCau > 0 ? formatLitersDisplay(o.DauCau) : ""
            };
        });
    }

    function rowMatchesWarehouse(raw, needle) {
        const n = needle.trim().toLowerCase();
        if (!n) return true;
        try {
            return JSON.stringify(raw).toLowerCase().includes(n);
        } catch (e) {
            return true;
        }
    }

    function getNlFilters() {
        const fromRaw = document.getElementById("nl-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("nl-filter-to-date")?.value?.trim() ?? "";
        return {
            projectsSelected: getNlSelectedCongTrinhValues(),
            fromDate: dateKeyYmd(parseDateFlexible(fromRaw)),
            toDate: dateKeyYmd(parseDateFlexible(toRaw))
        };
    }

    /** Đồng bộ dòng tiêu đề kỳ với đúng hai ô «Từ ngày» / «Đến ngày» (gõ tay hoặc lịch). */
    function syncNlDateRangeBanner() {
        const dr = document.getElementById("nl-date-range");
        if (!dr) return;
        const fromRaw = document.getElementById("nl-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("nl-filter-to-date")?.value?.trim() ?? "";
        const dFrom = fromRaw ? parseDateFlexible(fromRaw) : null;
        const dTo = toRaw ? parseDateFlexible(toRaw) : null;
        const labelFrom =
            !fromRaw ? "…" : dFrom && !isNaN(dFrom.getTime()) ? formatDateVn(dFrom) : fromRaw;
        const labelTo = !toRaw ? "…" : dTo && !isNaN(dTo.getTime()) ? formatDateVn(dTo) : toRaw;
        dr.textContent = `(Từ ngày ${labelFrom} tới ngày ${labelTo})`;
    }

    function fuelDisplay(v) {
        const s = String(v ?? "").trim();
        if (s === "" || s === "-") return "-";
        return s;
    }

    function renderNlTable(rows, tonDauKyFooter, nhapTrongThangFooter) {
        const tbody = document.getElementById("nl-tbody");
        if (!tbody) return;

        function cellTonDau(k) {
            if (!tonDauKyFooter?.hasAny) return "-";
            const v = tonDauKyFooter[k];
            if (v == null || isNaN(v)) return "-";
            return formatLitersDisplay(v);
        }

        const data = rows || [];
        if (!data.length) {
            tbody.innerHTML =
                '<tr><td colspan="9" class="text-center" style="padding:12px;">Không có dữ liệu «Nhập xuất luân chuyển CT» phù hợp bộ lọc.</td></tr>';
            return;
        }

        let html = "";
        let stt = 0;

        for (const row of data) {
            stt += 1;
            const ten = escapeHtml(String(row.TenThietBi ?? ""));
            const nguoiPt = escapeHtml(String(row.NguoiPhuTrach ?? ""));
            html += `<tr>
                <td class="text-center">${stt}</td>
                <td class="text-left">${ten}</td>
                <td class="text-left">${nguoiPt}</td>
                <td class="text-center">${escapeHtml(fuelDisplay(row.LuongDauTieuHao))}</td>
                <td class="text-center">${escapeHtml(fuelDisplay(row.TongMo))}</td>
                <td class="text-center">${escapeHtml(fuelDisplay(row.TongNhot))}</td>
                <td class="text-center">${escapeHtml(fuelDisplay(row.DauThuyLuc))}</td>
                <td class="text-center">${escapeHtml(fuelDisplay(row.DauCau))}</td>
                <td class="text-left"></td>
            </tr>`;
        }

        const leaf = data;
        function sumColNumber(key) {
            let s = 0;
            let any = false;
            for (const r of leaf) {
                const q = parseNumericForLiters(r[key]);
                if (q != null && !isNaN(q)) {
                    s += q;
                    any = true;
                }
            }
            return any ? s : null;
        }
        const xuat = {
            LuongDauTieuHao: sumColNumber("LuongDauTieuHao"),
            TongMo: sumColNumber("TongMo"),
            TongNhot: sumColNumber("TongNhot"),
            DauThuyLuc: sumColNumber("DauThuyLuc"),
            DauCau: sumColNumber("DauCau")
        };
        function displayNumberOrDash(n) {
            return n == null || isNaN(n) ? "-" : formatLitersDisplay(n);
        }
        function displayFooterByKey(footer, key) {
            if (!footer?.hasAny) return "-";
            const v = footer[key];
            if (v == null || isNaN(v)) return "-";
            return formatLitersDisplay(v);
        }
        function closingCellDisplay(key) {
            const hasB = tonDauKyFooter?.hasAny && tonDauKyFooter[key] != null && !isNaN(tonDauKyFooter[key]);
            const hasC = nhapTrongThangFooter?.hasAny && nhapTrongThangFooter[key] != null && !isNaN(nhapTrongThangFooter[key]);
            const hasA = xuat[key] != null && !isNaN(xuat[key]);
            if (!hasB && !hasC && !hasA) return "-";
            const b = hasB ? tonDauKyFooter[key] : 0;
            const c = hasC ? nhapTrongThangFooter[key] : 0;
            const a = hasA ? xuat[key] : 0;
            return formatLitersDisplay(b + c - a);
        }

        html += `<tr class="bg-teal">
            <td></td><td class="text-left">Tổng xuất</td><td></td>
            <td class="text-center">${escapeHtml(displayNumberOrDash(xuat.LuongDauTieuHao))}</td>
            <td class="text-center">${escapeHtml(displayNumberOrDash(xuat.TongMo))}</td>
            <td class="text-center">${escapeHtml(displayNumberOrDash(xuat.TongNhot))}</td>
            <td class="text-center">${escapeHtml(displayNumberOrDash(xuat.DauThuyLuc))}</td>
            <td class="text-center">${escapeHtml(displayNumberOrDash(xuat.DauCau))}</td>
            <td class="text-center">A</td>
        </tr>`;
        html += `<tr class="bg-teal">
            <td></td><td class="text-left">Tồn đầu kỳ</td><td></td>
            <td class="text-center">${escapeHtml(cellTonDau("LuongDauTieuHao"))}</td>
            <td class="text-center">${escapeHtml(cellTonDau("TongMo"))}</td>
            <td class="text-center">${escapeHtml(cellTonDau("TongNhot"))}</td>
            <td class="text-center">${escapeHtml(cellTonDau("DauThuyLuc"))}</td>
            <td class="text-center">${escapeHtml(cellTonDau("DauCau"))}</td>
            <td class="text-center">B</td>
        </tr>`;
        html += `<tr class="bg-teal">
            <td></td><td class="text-left">Nhập trong tháng</td><td></td>
            <td class="text-center">${escapeHtml(displayFooterByKey(nhapTrongThangFooter, "LuongDauTieuHao"))}</td>
            <td class="text-center">${escapeHtml(displayFooterByKey(nhapTrongThangFooter, "TongMo"))}</td>
            <td class="text-center">${escapeHtml(displayFooterByKey(nhapTrongThangFooter, "TongNhot"))}</td>
            <td class="text-center">${escapeHtml(displayFooterByKey(nhapTrongThangFooter, "DauThuyLuc"))}</td>
            <td class="text-center">${escapeHtml(displayFooterByKey(nhapTrongThangFooter, "DauCau"))}</td>
            <td class="text-center">C</td>
        </tr>`;
        html += `<tr class="bg-teal">
            <td></td><td class="text-left">Tồn cuối tháng</td><td></td>
            <td class="text-center">${escapeHtml(closingCellDisplay("LuongDauTieuHao"))}</td>
            <td class="text-center">${escapeHtml(closingCellDisplay("TongMo"))}</td>
            <td class="text-center">${escapeHtml(closingCellDisplay("TongNhot"))}</td>
            <td class="text-center">${escapeHtml(closingCellDisplay("DauThuyLuc"))}</td>
            <td class="text-center">${escapeHtml(closingCellDisplay("DauCau"))}</td>
            <td class="text-center">D=B+C-A</td>
        </tr>`;

        tbody.innerHTML = html;
    }

    function buildDateRangeFromFilters(filters) {
        let from = parseDateInputYmd(filters?.fromDate);
        let to = parseDateInputYmd(filters?.toDate);
        if (!from && !to) return null;
        if (!from) from = to;
        if (!to) to = from;
        let start = new Date(from.getFullYear(), from.getMonth(), from.getDate(), 0, 0, 0, 0);
        let end = new Date(to.getFullYear(), to.getMonth(), to.getDate(), 23, 59, 59, 999);
        if (start > end) {
            const tmp = start;
            start = end;
            end = tmp;
        }
        return { start, end };
    }

    function rowDateInRange(dateValue, dateRange) {
        if (!dateRange) return true;
        if (!dateValue || isNaN(dateValue.getTime())) return false;
        return dateValue >= dateRange.start && dateValue <= dateRange.end;
    }

    function rowMatchesSelectedProjectsByRawColumns(row, selectedProjects) {
        if (!selectedProjects?.length) return true;
        for (const p of selectedProjects) {
            if (rawCellMatchesNeedle(row, "project", p)) return true;
        }
        return false;
    }

    function rowMatchesSelectedCongTrinhLabel(row, selectedProjects) {
        if (!selectedProjects?.length) return true;
        const label = pickNtCtTenNoiQuanLy(row);
        if (!label) return false;
        return selectedProjects.some((p) => p === label || taiSanNamesLooselyEqual(p, label));
    }

    function switchNlLoadedTab(which) {
        document.querySelectorAll(".nl-loaded-tab").forEach((b) => {
            b.classList.toggle("active", b.getAttribute("data-nl-prev") === which);
        });
        document.querySelectorAll(".nl-loaded-panel").forEach((p) => {
            p.classList.toggle("is-visible", p.id === `nl-loaded-panel-${which}`);
        });
    }

    function closeNlLoadedDataModal() {
        const modal = document.getElementById("nl-loaded-data-modal");
        if (!modal) return;
        modal.style.display = "none";
        modal.setAttribute("aria-hidden", "true");
        document.body.style.overflow = "";
    }

    function openNlLoadedDataModal() {
        const modal = document.getElementById("nl-loaded-data-modal");
        if (!modal) return;
        if (cacheNxlcCtRows == null) {
            alert("Chưa tải dữ liệu AppSheet. Bấm «Tải dữ liệu AppSheet» trước.");
            return;
        }

        const filters = getNlFilters();
        const dateRange = buildDateRangeFromFilters(filters);
        const selectedProjects = filters.projectsSelected || [];

        const nxlcAll = cacheNxlcCtRows ?? [];
        const nxlcView = nxlcAll.filter(
            (r) => rowDateInRange(parseNxlcCtNgayCell(r), dateRange) && rowMatchesSelectedProjectsByRawColumns(r, selectedProjects)
        );

        const ntCtAll = cacheNtCtRows ?? [];
        const ntCtDateFiltered = filterCtRowsByNgayLamViec(ntCtAll, filters);
        const ntCtView = ntCtDateFiltered.filter((r) => rowMatchesSelectedCongTrinhLabel(r, selectedProjects));

        const ntNlAll = cacheNtNlRows ?? [];
        const ntNlView = ntNlAll.filter((r) => rowDateInRange(parseNlNgayCell(r), dateRange));

        const dkAll = cacheTonKhoDkRows ?? [];
        const dkView = dkAll.filter((r) => tonKhoDkRowDateInRange(r, dateRange));

        const dsAll = cacheDsTaiSanRows ?? [];
        const dsView = dsAll;

        const nxEl = document.getElementById("nl-loaded-panel-nxlc");
        const dkEl = document.getElementById("nl-loaded-panel-dk");
        const ctEl = document.getElementById("nl-loaded-panel-ntct");
        const nlEl = document.getElementById("nl-loaded-panel-ntnl");
        const dsEl = document.getElementById("nl-loaded-panel-ds");
        if (nxEl) nxEl.innerHTML = buildRawPreviewTableHtml(nxlcView);
        if (dkEl) dkEl.innerHTML = buildRawPreviewTableHtml(dkView);
        if (ctEl) ctEl.innerHTML = buildRawPreviewTableHtml(ntCtView);
        if (nlEl) nlEl.innerHTML = buildRawPreviewTableHtml(ntNlView);
        if (dsEl) dsEl.innerHTML = buildRawPreviewTableHtml(dsView);

        const tNx = document.getElementById("nl-tab-nxlc");
        const tDk = document.getElementById("nl-tab-dk");
        const tCt = document.getElementById("nl-tab-ntct");
        const tNl = document.getElementById("nl-tab-ntnl");
        const tDs = document.getElementById("nl-tab-ds");
        if (tNx) tNx.textContent = `Nhập xuất luân chuyển CT (${formatPreviewTabCount(nxlcView.length, nxlcAll.length)})`;
        if (tDk) tDk.textContent = `Tồn kho ĐK (${formatPreviewTabCount(dkView.length, dkAll.length)})`;
        if (tCt) tCt.textContent = `Nhật trình máy CT (${formatPreviewTabCount(ntCtView.length, ntCtAll.length)})`;
        if (tNl) tNl.textContent = `Nhật trình NL (${formatPreviewTabCount(ntNlView.length, ntNlAll.length)})`;
        if (tDs) tDs.textContent = `Danh sách tài sản (${formatPreviewTabCount(dsView.length, dsAll.length)})`;

        switchNlLoadedTab("nxlc");
        modal.style.display = "flex";
        modal.setAttribute("aria-hidden", "false");
        document.body.style.overflow = "hidden";
    }

    async function loadFromAppSheet(forceRefresh) {
        const btn = document.getElementById("nl-btn-load");
        if (btn) btn.disabled = true;

        try {
            const filters = getNlFilters();

            if (forceRefresh) {
                cacheNxlcCtRows = null;
                cacheDsTaiSanRows = null;
                cacheNtCtRows = null;
                cacheNtNlRows = null;
                cacheTonKhoDkRows = null;
            }

            if (!cacheNxlcCtRows) {
                cacheNxlcCtRows = await fetchAppSheetTable(TABLE_NXLC_CT);
            }
            if (!cacheDsTaiSanRows) {
                cacheDsTaiSanRows = await fetchAppSheetTable(TABLE_DS_TAI_SAN);
            }
            if (!cacheNtCtRows) {
                cacheNtCtRows = await fetchAppSheetTable(TABLE_NT_CT);
            }
            if (!cacheNtNlRows) {
                cacheNtNlRows = await fetchAppSheetTable(TABLE_NT_NL);
            }
            if (!cacheTonKhoDkRows) {
                cacheTonKhoDkRows = await fetchAppSheetTable(TABLE_TON_KHO_DK);
            }
            populateNlCongTrinhCheckboxList(cacheNtCtRows);

            const tenLmByKey = buildTenLmByTaiSanFromDanhSachTaiSan(cacheDsTaiSanRows);
            const tenNoiQuanLyByKey = buildTenNoiQuanLyByTaiSanFromNtCt(cacheNtCtRows);

            let normalized = buildAggregatedRowsFromNxlcCt(cacheNxlcCtRows, filters);
            normalized = normalized.map((row) => ({
                ...row,
                CongTrinh: resolveTenNoiQuanLyForTenThietBi(
                    row.TenThietBi,
                    tenNoiQuanLyByKey,
                    cacheNtCtRows,
                    cacheDsTaiSanRows
                ),
                NguoiPhuTrach: resolveTenLmForTenXmtb(row.TenThietBi, tenLmByKey, cacheDsTaiSanRows)
            }));
            normalized = filterNlRowsBySelectedCongTrinh(normalized, filters.projectsSelected);
            const tonDauKyFooter = buildTonDauKyTotalsFromTonKhoDk(cacheTonKhoDkRows, filters);
            const nhapTrongThangFooter = buildNhapTrongThangTotalsFromNxlcCt(normalized, cacheNxlcCtRows, filters);
            renderNlTable(normalized, tonDauKyFooter, nhapTrongThangFooter);

            syncNlDateRangeBanner();
            syncNlCongTrinhTitleBanner();
        } catch (e) {
            console.error(e);
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function initNlAppSheetUi() {
        wireNativeDatePickerButton("nl-filter-from-date", "nl-filter-from-date-picker", "nl-filter-from-date-cal");
        wireNativeDatePickerButton("nl-filter-to-date", "nl-filter-to-date-picker", "nl-filter-to-date-cal");

        for (const id of ["nl-filter-from-date", "nl-filter-to-date"]) {
            const el = document.getElementById(id);
            if (!el) continue;
            const onDateFilterChange = () => syncNlDateRangeBanner();
            el.addEventListener("input", onDateFilterChange);
            el.addEventListener("keyup", onDateFilterChange);
            el.addEventListener("change", onDateFilterChange);
            el.addEventListener("blur", onDateFilterChange);
            el.addEventListener("paste", () => setTimeout(onDateFilterChange, 0));
        }
        syncNlDateRangeBanner();

        const btn = document.getElementById("nl-btn-load");
        const btnRef = document.getElementById("nl-btn-refresh");
        if (btn) btn.addEventListener("click", () => loadFromAppSheet(false));
        if (btnRef) btnRef.addEventListener("click", () => loadFromAppSheet(true));

        const panelCt = document.getElementById("nl-filter-cong-trinh-panel");
        const btnSelAll = document.getElementById("nl-cong-trinh-select-all");
        const btnClear = document.getElementById("nl-cong-trinh-clear");
        if (btnSelAll && panelCt) {
            btnSelAll.addEventListener("click", () => {
                for (const el of panelCt.querySelectorAll('input[type="checkbox"][name="nl-cong-trinh"]')) {
                    el.checked = true;
                }
                syncNlCongTrinhTitleBanner();
            });
        }
        if (btnClear && panelCt) {
            btnClear.addEventListener("click", () => {
                for (const el of panelCt.querySelectorAll('input[type="checkbox"][name="nl-cong-trinh"]')) {
                    el.checked = false;
                }
                syncNlCongTrinhTitleBanner();
            });
        }
        panelCt?.addEventListener("change", (e) => {
            if (e.target?.matches?.('input[type="checkbox"][name="nl-cong-trinh"]')) syncNlCongTrinhTitleBanner();
        });

        const search = document.getElementById("nl-search");
        if (search) {
            search.addEventListener("input", () => {
                const q = search.value.trim().toLowerCase();
                const tbody = document.getElementById("nl-tbody");
                if (!tbody) return;
                for (const tr of tbody.querySelectorAll("tr")) {
                    const t = tr.textContent.toLowerCase();
                    tr.style.display = !q || t.includes(q) ? "" : "none";
                }
            });
        }

        document.getElementById("nl-btn-view-loaded-tables")?.addEventListener("click", openNlLoadedDataModal);
        document.getElementById("nl-loaded-close")?.addEventListener("click", closeNlLoadedDataModal);
        document.getElementById("nl-loaded-data-modal")?.addEventListener("click", (e) => {
            if (e.target && e.target.id === "nl-loaded-data-modal") closeNlLoadedDataModal();
        });
        document.querySelectorAll(".nl-loaded-tab").forEach((btnTab) => {
            btnTab.addEventListener("click", () => switchNlLoadedTab(btnTab.getAttribute("data-nl-prev") || "nxlc"));
        });
        document.addEventListener("keydown", (e) => {
            if (e.key !== "Escape") return;
            const modal = document.getElementById("nl-loaded-data-modal");
            if (!modal || modal.getAttribute("aria-hidden") !== "false") return;
            closeNlLoadedDataModal();
        });

        loadFromAppSheet(false);
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", initNlAppSheetUi);
    } else {
        initNlAppSheetUi();
    }
})();
