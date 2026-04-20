/**
 * Báo cáo tổng hợp xe máy thiết bị theo nơi quản lý — AppSheet.
 * Nguồn: «Danh sách tài sản», gom nhóm theo cột «Nơi quản lý» (và tên cột tương đương).
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
    const TBX_PREVIEW_ROW_CAP = 500;

    let cacheDsRows = null;
    let cacheTonKhoDkRows = null;
    let cacheNxlcCtRows = null;

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

    function collectPreviewKeys(rows) {
        const set = new Set();
        for (const r of rows || []) {
            if (!r || typeof r !== "object") continue;
            for (const k of Object.keys(r)) set.add(k);
        }
        return [...set].sort((a, b) => String(a).localeCompare(String(b), "vi"));
    }

    function previewCellValue(v) {
        if (v == null) return "";
        const s = cellDisplayString(v).trim();
        if (s) return s.length > 800 ? `${s.slice(0, 800)}...` : s;
        try {
            const j = JSON.stringify(v);
            if (j && j !== "{}" && j !== "null") return j.length > 800 ? `${j.slice(0, 800)}...` : j;
        } catch (_) {}
        return String(v);
    }

    function buildRawPreviewTableHtml(rows) {
        const all = rows || [];
        if (!all.length) return '<p class="tbx-prev-note">Không có dòng.</p>';
        const total = all.length;
        const show = all.slice(0, TBX_PREVIEW_ROW_CAP);
        const keys = collectPreviewKeys(all);
        if (!keys.length) return `<p class="tbx-prev-note">${total} dòng - không có khóa cột.</p>`;
        const thead = `<tr>${keys.map((k) => `<th>${escapeHtml(k)}</th>`).join("")}</tr>`;
        const tbody = show
            .map((row) => `<tr>${keys.map((k) => `<td>${escapeHtml(previewCellValue(row[k]))}</td>`).join("")}</tr>`)
            .join("");
        const note =
            (total > TBX_PREVIEW_ROW_CAP ? `Hiển thị tối đa ${TBX_PREVIEW_ROW_CAP}/${total} dòng. ` : `${total} dòng. `) +
            "Kéo ngang nếu bảng rộng.";
        return `<p class="tbx-prev-note">${escapeHtml(note)}</p><div class="tbx-prev-scroll"><table class="tbx-prev-table"><thead>${thead}</thead><tbody>${tbody}</tbody></table></div>`;
    }

    function closeLoadedDataModal() {
        const modal = document.getElementById("tbx-loaded-data-modal");
        if (!modal) return;
        modal.style.display = "none";
        modal.setAttribute("aria-hidden", "true");
        document.body.style.overflow = "";
    }

    function openLoadedDataModal() {
        const modal = document.getElementById("tbx-loaded-data-modal");
        const panel = document.getElementById("tbx-loaded-panel");
        if (!modal || !panel) return;
        if (cacheDsRows == null) {
            alert("Chưa tải dữ liệu AppSheet. Bấm «Tải dữ liệu AppSheet» trước.");
            return;
        }
        const filters = getFilters();
        const view = (cacheDsRows || []).filter((r) => rowMatchesFilters(r, filters));
        panel.innerHTML = buildRawPreviewTableHtml(view);
        modal.style.display = "flex";
        modal.setAttribute("aria-hidden", "false");
        document.body.style.overflow = "hidden";
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

    function pickDsNoiQuanLy(row) {
        const keys = [
            "Nơi quản lý",
            "Noi quan ly",
            "Tên nơi quản lý",
            "Ten noi quan ly",
            "Nhóm nơi quản lý",
            "Nhom noi quan ly",
            "Tên công trình",
            "Ten cong trinh",
            "Công trường",
            "Cong truong"
        ];
        const direct = pickFirstCell(row, keys);
        if (direct) return direct;
        for (const [k, v] of Object.entries(row || {})) {
            const kn = String(k).toLowerCase();
            if (/n[oơ]i.*qu[aả]n|noi.*quan|cong.*trinh|cong.*truong/i.test(kn)) {
                const s = cellDisplayString(v).trim();
                if (s) return s;
            }
        }
        return "";
    }

    /** Giá trị hiển thị / lọc «Nơi thi công»: ưu tiên cột riêng, không có thì dùng «Nơi quản lý» (cùng nhóm báo cáo). */
    function pickDsNoiThiCongForFilter(row) {
        const strict = pickFirstCell(row, [
            "Nơi thi công",
            "Noi thi cong",
            "Địa điểm thi công",
            "Dia diem thi cong",
            "Khu vực thi công"
        ]).trim();
        if (strict) return strict;
        return pickDsNoiQuanLy(row).trim();
    }

    /** Chỉ các cột «Công trình» / dự án — không trùng fallback nơi quản lý. */
    function pickDsCongTrinhForFilter(row) {
        return pickFirstCell(row, [
            "Công trình",
            "Tên công trình",
            "Ten cong trinh",
            "Dự án",
            "Du an",
            "Tên dự án",
            "Ten du an"
        ]).trim();
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

    function pickDsGhiChu(row) {
        return pickFirstCell(row, ["Ghi chú", "Ghi chu", "GhiChu", "Mô tả", "Mo ta"]);
    }

    function normalizeNameKey(s) {
        return String(s ?? "")
            .trim()
            .normalize("NFD")
            .replace(/\u0300-\u036f/g, "")
            .toLowerCase()
            .replace(/[^a-z0-9]/g, "");
    }

    function parseTonKhoDkTenNlCell(row) {
        const keys = [
            "Tên NL",
            "Ten NL",
            "Tên nhiên liệu",
            "Ten nhien lieu",
            "Loại nhiên liệu",
            "Loai nhien lieu",
            "Mã NL",
            "Ma NL"
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
            if (!/ten\s*nl|ten\s*nhien\s*lieu|loai\s*nhien\s*lieu|ma\s*nl/.test(kn)) continue;
            const t = cellDisplayString(v).trim();
            if (t) return t;
        }
        return "";
    }

    function parseTonKhoDkSlTonCell(row) {
        const keys = ["Sl tồn ĐK", "SL tồn ĐK", "Sl tồn DK", "SL tồn DK", "SL ton DK", "Sl ton DK"];
        for (const k of keys) {
            const q = parseNumeric(row?.[k]);
            if (q != null && !isNaN(q)) return q;
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (!/sl\s*t[oô]n\s*[đd]k|sl\s*ton\s*dk/i.test(String(k))) continue;
            const q = parseNumeric(v);
            if (q != null && !isNaN(q)) return q;
        }
        return null;
    }

    function getNxlcTenNhienLieuCell(row) {
        const keys = ["Tên nhiên liệu", "Ten nhien lieu", "Tên NL", "Ten NL", "Loại nhiên liệu", "Loai nhien lieu"];
        for (const k of keys) {
            const t = cellDisplayString(row?.[k]).trim();
            if (t) return t;
        }
        for (const [k, v] of Object.entries(row || {})) {
            const kn = String(k)
                .normalize("NFD")
                .replace(/\u0300-\u036f/g, "")
                .toLowerCase();
            if (!/ten\s*nhien\s*lieu|ten\s*nl|loai\s*nhien\s*lieu/.test(kn)) continue;
            const t = cellDisplayString(v).trim();
            if (t) return t;
        }
        return "";
    }

    function parseNxlcSoLuongNlCell(row) {
        const keys = [
            "Số lượng NL",
            "So luong NL",
            "Số lượng nhiên liệu",
            "So luong nhien lieu",
            "SL NL",
            "SL",
            "Sl"
        ];
        for (const k of keys) {
            const q = parseNumeric(row?.[k]);
            if (q != null && !isNaN(q)) return q;
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (!/so\s*luong.*nl|s[oố]\s*l[ưu]ợ?ng.*nl|sl.*nl/i.test(String(k))) continue;
            const q = parseNumeric(v);
            if (q != null && !isNaN(q)) return q;
        }
        return null;
    }

    function parseNxlcThanhTienNlCell(row) {
        const keys = ["Thành tiền NL", "Thanh tien NL", "Thành tiền nhiên liệu", "Thanh tien nhien lieu", "Tiền NL"];
        for (const k of keys) {
            const q = parseNumeric(row?.[k]);
            if (q != null && !isNaN(q)) return q;
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (!/thanh\s*tien.*nl|th[aà]nh\s*ti[eề]n.*nl|tien.*nl/i.test(String(k))) continue;
            const q = parseNumeric(v);
            if (q != null && !isNaN(q)) return q;
        }
        return null;
    }

    function loaiPhieuIsNhap(rawVal) {
        const n = String(rawVal ?? "")
            .trim()
            .toLowerCase()
            .normalize("NFD")
            .replace(/\u0300-\u036f/g, "");
        if (!n) return false;
        const compact = n.replace(/[^a-z0-9]/g, "");
        if (compact === "nhap" || compact === "nhp") return true;
        if (compact.startsWith("nhap")) return true;
        if (compact.includes("phieunhap") || compact.includes("nhapkho") || compact.includes("nhaphang")) return true;
        return false;
    }

    function loaiPhieuIsXuat(rawVal) {
        const n = String(rawVal ?? "")
            .trim()
            .toLowerCase()
            .normalize("NFD")
            .replace(/\u0300-\u036f/g, "");
        if (!n) return false;
        const compact = n.replace(/[^a-z0-9]/g, "");
        if (compact === "xuat" || compact === "xut") return true;
        if (compact.startsWith("xuat")) return true;
        if (compact.includes("phieuxuat") || compact.includes("xuatkho") || compact.includes("xuathang")) return true;
        return false;
    }

    function getNxlcLoaiPhieuCell(row) {
        const keys = ["Loại phiếu", "Loai phieu", "Loại phiếu NX", "Loai phieu NX", "Kiểu phiếu", "Kieu phieu"];
        for (const k of keys) {
            const t = cellDisplayString(row?.[k]).trim();
            if (t) return t;
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (v == null || String(v).trim() === "") continue;
            if (/lo[aạ]i\s*phi[ếe]u|loai\s*phieu/i.test(String(k))) return v;
        }
        for (const v of Object.values(row || {})) {
            if (v == null || String(v).trim() === "") continue;
            if (loaiPhieuIsNhap(v) || loaiPhieuIsXuat(v)) return v;
        }
        return "";
    }

    /**
     * Theo yêu cầu nghiệp vụ:
     * «NHẬP TRONG KỲ» lấy từ «Nhập xuất luân chuyển CT» với «Loại phiếu = Nhập»,
     * khớp «Tên nhiên liệu» với «TÊN THIẾT BỊ».
     * - «Số lượng NL» -> cột SL
     * - «Thành tiền NL» -> cột GIÁ TRỊ
     */
    function buildNhapTrongKyByTenThietBiFromNxlcCt(nxlcCtRows) {
        const map = new Map();
        for (const r of nxlcCtRows || []) {
            const loaiPhieu = getNxlcLoaiPhieuCell(r);
            if (!loaiPhieuIsNhap(loaiPhieu)) continue;
            const tenNl = getNxlcTenNhienLieuCell(r);
            const key = normalizeNameKey(tenNl);
            if (!key) continue;
            const sl = parseNxlcSoLuongNlCell(r);
            const gt = parseNxlcThanhTienNlCell(r);
            if ((sl == null || isNaN(sl)) && (gt == null || isNaN(gt))) continue;
            if (!map.has(key)) map.set(key, { nhapSl: 0, nhapGt: 0 });
            const cur = map.get(key);
            if (sl != null && !isNaN(sl)) cur.nhapSl += sl;
            if (gt != null && !isNaN(gt)) cur.nhapGt += gt;
        }
        return map;
    }

    /**
     * Theo yêu cầu nghiệp vụ:
     * «XUẤT TRONG KỲ» lấy từ «Nhập xuất luân chuyển CT» với «Loại phiếu = Xuất»,
     * khớp «Tên nhiên liệu» với «TÊN THIẾT BỊ».
     * - «Số lượng NL» -> cột SL
     * - «Thành tiền NL» -> cột GIÁ TRỊ
     */
    function buildXuatTrongKyByTenThietBiFromNxlcCt(nxlcCtRows) {
        const map = new Map();
        for (const r of nxlcCtRows || []) {
            const loaiPhieu = getNxlcLoaiPhieuCell(r);
            if (!loaiPhieuIsXuat(loaiPhieu)) continue;
            const tenNl = getNxlcTenNhienLieuCell(r);
            const key = normalizeNameKey(tenNl);
            if (!key) continue;
            const sl = parseNxlcSoLuongNlCell(r);
            const gt = parseNxlcThanhTienNlCell(r);
            if ((sl == null || isNaN(sl)) && (gt == null || isNaN(gt))) continue;
            if (!map.has(key)) map.set(key, { xuatSl: 0, xuatGt: 0 });
            const cur = map.get(key);
            if (sl != null && !isNaN(sl)) cur.xuatSl += sl;
            if (gt != null && !isNaN(gt)) cur.xuatGt += gt;
        }
        return map;
    }

    /**
     * Theo yêu cầu nghiệp vụ:
     * «TỒN ĐẦU KỲ» cột SL lấy từ bảng «Tồn kho ĐK»,
     * khớp «Tên NL» với «TÊN THIẾT BỊ», và lấy tổng «Sl tồn ĐK».
     */
    function buildTonDauSlByTenThietBiFromTonKhoDk(tonKhoDkRows) {
        const map = new Map();
        for (const r of tonKhoDkRows || []) {
            const tenNl = parseTonKhoDkTenNlCell(r);
            const key = normalizeNameKey(tenNl);
            if (!key) continue;
            const slTon = parseTonKhoDkSlTonCell(r);
            if (slTon == null || isNaN(slTon)) continue;
            map.set(key, (map.get(key) ?? 0) + slTon);
        }
        return map;
    }

    /** Đọc ô số theo danh sách tên cột ưu tiên; nếu không có thì quét tên cột khớp regex. */
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

    /**
     * Map các cột thường gặp trên DS cho báo cáo tồn/nhập/xuất (nếu AppSheet có đặt tên tương tự).
     */
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

    function displayCell(n) {
        if (n == null || isNaN(n)) return "";
        return formatNum(n);
    }

    function getCheckedFilterValues(listId) {
        const root = document.getElementById(listId);
        if (!root) return [];
        return [...root.querySelectorAll('input[type="checkbox"]:checked')].map((cb) => cb.value);
    }

    function rowMatchesFilters(row, filters) {
        if (filters.noiThiCong) {
            if (pickDsNoiThiCongForFilter(row) !== filters.noiThiCong) return false;
        }
        if (filters.congTrinh.length > 0) {
            const v = pickDsCongTrinhForFilter(row);
            if (!v || !filters.congTrinh.includes(v)) return false;
        }
        if (filters.nhomThietBi.length > 0) {
            const v = pickDsNhomThietBi(row).trim();
            if (!v || !filters.nhomThietBi.includes(v)) return false;
        }
        return true;
    }

    function getFilters() {
        return {
            noiThiCong: document.getElementById("tbx-filter-noi-thi-cong")?.value?.trim() ?? "",
            congTrinh: getCheckedFilterValues("tbx-filter-cong-trinh-list"),
            nhomThietBi: getCheckedFilterValues("tbx-filter-nhom-thiet-bi-list")
        };
    }

    function populateFilterCheckboxes(rows) {
        const selNoi = document.getElementById("tbx-filter-noi-thi-cong");
        const listCt = document.getElementById("tbx-filter-cong-trinh-list");
        const listNhom = document.getElementById("tbx-filter-nhom-thiet-bi-list");
        if (!selNoi || !listCt || !listNhom) return;

        const prevNoi = selNoi.value?.trim() ?? "";
        const prevCt = new Set(getCheckedFilterValues("tbx-filter-cong-trinh-list"));
        const prevNhom = new Set(getCheckedFilterValues("tbx-filter-nhom-thiet-bi-list"));

        const setNoi = new Set();
        const setCt = new Set();
        const setNhom = new Set();
        for (const r of rows || []) {
            const n = pickDsNoiThiCongForFilter(r);
            if (n) setNoi.add(n);
            const c = pickDsCongTrinhForFilter(r);
            if (c) setCt.add(c);
            const h = pickDsNhomThietBi(r).trim();
            if (h) setNhom.add(h);
        }

        const sortVi = (a, b) => a.localeCompare(b, "vi");

        selNoi.innerHTML = "";
        const o0 = document.createElement("option");
        o0.value = "";
        o0.textContent = "--- Tất cả nơi thi công ---";
        selNoi.appendChild(o0);
        for (const v of [...setNoi].sort(sortVi)) {
            const o = document.createElement("option");
            o.value = v;
            o.textContent = v;
            selNoi.appendChild(o);
        }
        if (prevNoi && setNoi.has(prevNoi)) selNoi.value = prevNoi;
        else selNoi.value = "";

        function fillPanel(panel, valueSet, namePrefix, prevChecked) {
            panel.innerHTML = "";
            const sorted = [...valueSet].sort(sortVi);
            for (const v of sorted) {
                const label = document.createElement("label");
                const input = document.createElement("input");
                input.type = "checkbox";
                input.name = namePrefix;
                input.value = v;
                if (prevChecked.has(v)) input.checked = true;
                const span = document.createElement("span");
                span.textContent = v;
                label.appendChild(input);
                label.appendChild(span);
                panel.appendChild(label);
            }
        }

        fillPanel(listCt, setCt, "tbx-cong-trinh", prevCt);
        fillPanel(listNhom, setNhom, "tbx-nhom-thiet-bi", prevNhom);
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
        return buildStrictDate(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));
    }

    function parseDateFlexible(value) {
        if (value instanceof Date) return value;
        const ymd = parseDateInputYmd(value);
        if (ymd) return ymd;
        const s = String(value ?? "").trim();
        if (!s) return null;
        let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (m) {
            const a = parseInt(m[1], 10);
            const b = parseInt(m[2], 10);
            const y = parseInt(m[3], 10);
            if (a <= 12 && b > 12) return buildStrictDate(y, a, b);
            return buildStrictDate(y, b, a);
        }
        return null;
    }

    function formatDateVn(value) {
        const d = parseDateFlexible(value);
        if (!d || isNaN(d.getTime())) return "";
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        return `${dd}/${mm}/${d.getFullYear()}`;
    }

    function dateKeyYmd(d) {
        if (!d || isNaN(d.getTime())) return "";
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const dd = String(d.getDate()).padStart(2, "0");
        return `${yyyy}-${mm}-${dd}`;
    }

    function syncDateRangeBanner() {
        const el = document.getElementById("tbx-date-range");
        if (!el) return;
        const fromRaw = document.getElementById("tbx-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("tbx-filter-to-date")?.value?.trim() ?? "";
        const dFrom = fromRaw ? parseDateFlexible(fromRaw) : null;
        const dTo = toRaw ? parseDateFlexible(toRaw) : null;
        const lf = !fromRaw ? "…" : dFrom && !isNaN(dFrom.getTime()) ? formatDateVn(fromRaw) : fromRaw;
        const lt = !toRaw ? "…" : dTo && !isNaN(dTo.getTime()) ? formatDateVn(toRaw) : toRaw;
        el.textContent = `(Từ ngày ${lf} đến ngày ${lt})`;
    }

    function buildGroupedReport(dsRows, filters) {
        const filtered = (dsRows || []).filter((r) => rowMatchesFilters(r, filters));
        const byPlace = new Map();
        for (const r of filtered) {
            let place = pickDsNoiQuanLy(r).trim();
            if (!place) place = "(Chưa có nơi quản lý)";
            if (!byPlace.has(place)) byPlace.set(place, []);
            byPlace.get(place).push(r);
        }
        const groups = [...byPlace.entries()].sort((a, b) => a[0].localeCompare(b[0], "vi"));
        for (const [, rows] of groups) {
            rows.sort((a, b) => pickDsTenThietBi(a).localeCompare(pickDsTenThietBi(b), "vi"));
        }
        return groups;
    }

    function renderTable(groups, tonDauSlByTenThietBi, nhapTrongKyByTenThietBi, xuatTrongKyByTenThietBi) {
        const tbody = document.getElementById("tbx-tbody");
        if (!tbody) return;

        let letter = 64;
        let html = "";
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

        function calcTonCuoiByFormula(tonDau, nhap, xuat) {
            const a = tonDau != null && !isNaN(tonDau) ? tonDau : null;
            const b = nhap != null && !isNaN(nhap) ? nhap : null;
            const c = xuat != null && !isNaN(xuat) ? xuat : null;
            if (a == null && b == null && c == null) return null;
            return (a ?? 0) + (b ?? 0) - (c ?? 0);
        }

        function addTot(st) {
            for (const k of Object.keys(totals)) {
                const v = st[k];
                if (v != null && !isNaN(v)) {
                    totals[k] += v;
                    anyNum = true;
                }
            }
        }

        for (const [place, rows] of groups) {
            letter += 1;
            const ch = letter <= 90 ? String.fromCharCode(letter) : "•";
            html += `<tr class="bg-table-group group-row">
                <td>${escapeHtml(ch)}</td>
                <td colspan="12">${escapeHtml(place)}</td>
            </tr>`;

            for (const r of rows) {
                const st = extractStockCells(r);
                const ma = pickDsMaThietBi(r);
                const ten = pickDsTenThietBi(r);
                const tonKey = normalizeNameKey(ten);
                if (tonKey && tonDauSlByTenThietBi?.has(tonKey)) st.tonDauSl = tonDauSlByTenThietBi.get(tonKey);
                else st.tonDauSl = null;
                if (tonKey && nhapTrongKyByTenThietBi?.has(tonKey)) {
                    const nx = nhapTrongKyByTenThietBi.get(tonKey);
                    st.nhapSl = nx?.nhapSl ?? null;
                    st.nhapGt = nx?.nhapGt ?? null;
                } else {
                    st.nhapSl = null;
                    st.nhapGt = null;
                }
                if (tonKey && xuatTrongKyByTenThietBi?.has(tonKey)) {
                    const xx = xuatTrongKyByTenThietBi.get(tonKey);
                    st.xuatSl = xx?.xuatSl ?? null;
                    st.xuatGt = xx?.xuatGt ?? null;
                } else {
                    st.xuatSl = null;
                    st.xuatGt = null;
                }
                // Theo yêu cầu nghiệp vụ: Tồn cuối kỳ = Tồn đầu kỳ + Nhập trong kỳ - Xuất trong kỳ.
                st.tonCuoiSl = calcTonCuoiByFormula(st.tonDauSl, st.nhapSl, st.xuatSl);
                st.tonCuoiGt = calcTonCuoiByFormula(st.tonDauGt, st.nhapGt, st.xuatGt);
                addTot(st);
                const dvt = pickDsDvt(r);
                const nhom = pickDsNhomThietBi(r);
                const gc = pickDsGhiChu(r);
                html += `<tr class="data-row">
                    <td class="text-center">${escapeHtml(ma)}</td>
                    <td class="pl-2 text-left">${escapeHtml(ten ? `- ${ten}` : "")}</td>
                    <td class="text-center">${escapeHtml(dvt)}</td>
                    <td class="text-center">${escapeHtml(displayCell(st.tonDauSl))}</td>
                    <td class="text-right pr-1">${escapeHtml(displayCell(st.tonDauGt))}</td>
                    <td class="text-center">${escapeHtml(displayCell(st.nhapSl))}</td>
                    <td class="text-right pr-1">${escapeHtml(displayCell(st.nhapGt))}</td>
                    <td class="text-center">${escapeHtml(displayCell(st.xuatSl))}</td>
                    <td class="text-right pr-1">${escapeHtml(displayCell(st.xuatGt))}</td>
                    <td class="text-center">${escapeHtml(displayCell(st.tonCuoiSl))}</td>
                    <td class="text-right pr-1">${escapeHtml(displayCell(st.tonCuoiGt))}</td>
                    <td class="text-center">${escapeHtml(nhom)}</td>
                    <td class="text-left pl-1">${escapeHtml(gc)}</td>
                </tr>`;
            }
        }

        if (!groups.length) {
            tbody.innerHTML =
                '<tr><td colspan="13" class="text-center py-4 text-gray-500">Không có dòng «Danh sách tài sản» phù hợp bộ lọc.</td></tr>';
            return;
        }

        const tf = (k) => (anyNum ? formatNum(totals[k]) : "");
        html += `<tr class="bg-table-header footer-row">
            <td colspan="3" class="text-center">TỔNG CỘNG</td>
            <td class="text-center">${escapeHtml(tf("tonDauSl"))}</td>
            <td class="text-right pr-1">${escapeHtml(tf("tonDauGt"))}</td>
            <td class="text-center">${escapeHtml(tf("nhapSl"))}</td>
            <td class="text-right pr-1">${escapeHtml(tf("nhapGt"))}</td>
            <td class="text-center">${escapeHtml(tf("xuatSl"))}</td>
            <td class="text-right pr-1">${escapeHtml(tf("xuatGt"))}</td>
            <td class="text-center">${escapeHtml(tf("tonCuoiSl"))}</td>
            <td class="text-right pr-1">${escapeHtml(tf("tonCuoiGt"))}</td>
            <td colspan="2"></td>
        </tr>`;

        tbody.innerHTML = html;
    }

    function setStatus(msg, isErr) {
        const el = document.getElementById("tbx-appsheet-status");
        if (!el) return;
        el.textContent = msg;
        el.classList.toggle("text-red-600", !!isErr);
        el.classList.toggle("text-header-green", !isErr);
    }

    async function loadFromAppSheet(forceRefresh) {
        const btn = document.getElementById("tbx-btn-load");
        if (btn) btn.disabled = true;
        setStatus("Đang tải AppSheet…", false);
        try {
            if (forceRefresh) {
                cacheDsRows = null;
                cacheTonKhoDkRows = null;
                cacheNxlcCtRows = null;
            }
            if (!cacheDsRows) {
                setStatus(`Đang tải «${TABLE_DS_TAI_SAN}»…`, false);
                cacheDsRows = await fetchAppSheetTable(TABLE_DS_TAI_SAN);
            }
            if (!cacheTonKhoDkRows) {
                setStatus(`Đang tải «${TABLE_TON_KHO_DK}»…`, false);
                cacheTonKhoDkRows = await fetchAppSheetTable(TABLE_TON_KHO_DK);
            }
            if (!cacheNxlcCtRows) {
                setStatus(`Đang tải «${TABLE_NXLC_CT}»…`, false);
                cacheNxlcCtRows = await fetchAppSheetTable(TABLE_NXLC_CT);
            }
            populateFilterCheckboxes(cacheDsRows);
            const filters = getFilters();
            const groups = buildGroupedReport(cacheDsRows, filters);
            const tonDauSlByTenThietBi = buildTonDauSlByTenThietBiFromTonKhoDk(cacheTonKhoDkRows);
            const nhapTrongKyByTenThietBi = buildNhapTrongKyByTenThietBiFromNxlcCt(cacheNxlcCtRows);
            const xuatTrongKyByTenThietBi = buildXuatTrongKyByTenThietBiFromNxlcCt(cacheNxlcCtRows);
            renderTable(groups, tonDauSlByTenThietBi, nhapTrongKyByTenThietBi, xuatTrongKyByTenThietBi);
            syncDateRangeBanner();
            const nDs = cacheDsRows?.length ?? 0;
            const nDk = cacheTonKhoDkRows?.length ?? 0;
            const nCt = cacheNxlcCtRows?.length ?? 0;
            const g = groups.length;
            setStatus(`Đã tải ${nDs} dòng «${TABLE_DS_TAI_SAN}», ${nDk} dòng «${TABLE_TON_KHO_DK}», ${nCt} dòng «${TABLE_NXLC_CT}» — ${g} nhóm nơi quản lý.`, false);
        } catch (e) {
            console.error(e);
            setStatus(`Lỗi: ${e.message}`, true);
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function init() {
        const loadBtn = document.getElementById("tbx-btn-load");
        const refBtn = document.getElementById("tbx-btn-refresh");
        if (loadBtn) loadBtn.addEventListener("click", () => loadFromAppSheet(false));
        if (refBtn) refBtn.addEventListener("click", () => loadFromAppSheet(true));
        document.getElementById("tbx-btn-view-loaded-tables")?.addEventListener("click", openLoadedDataModal);
        document.getElementById("tbx-loaded-close")?.addEventListener("click", closeLoadedDataModal);
        document.getElementById("tbx-loaded-data-modal")?.addEventListener("click", (e) => {
            if (e.target && e.target.id === "tbx-loaded-data-modal") closeLoadedDataModal();
        });
        document.addEventListener("keydown", (e) => {
            if (e.key !== "Escape") return;
            const modal = document.getElementById("tbx-loaded-data-modal");
            if (!modal || modal.getAttribute("aria-hidden") !== "false") return;
            closeLoadedDataModal();
        });

        for (const id of ["tbx-filter-from-date", "tbx-filter-to-date"]) {
            const el = document.getElementById(id);
            if (!el) continue;
            const fn = () => syncDateRangeBanner();
            el.addEventListener("input", fn);
            el.addEventListener("change", fn);
            el.addEventListener("blur", fn);
        }

        document.getElementById("tbx-filter-noi-thi-cong")?.addEventListener("change", () => loadFromAppSheet(false));

        const ctList = document.getElementById("tbx-filter-cong-trinh-list");
        if (ctList) ctList.addEventListener("change", () => loadFromAppSheet(false));
        document.getElementById("tbx-ct-check-all")?.addEventListener("click", () => {
            document.querySelectorAll("#tbx-filter-cong-trinh-list input[type='checkbox']").forEach((cb) => {
                cb.checked = true;
            });
            loadFromAppSheet(false);
        });
        document.getElementById("tbx-ct-check-none")?.addEventListener("click", () => {
            document.querySelectorAll("#tbx-filter-cong-trinh-list input[type='checkbox']").forEach((cb) => {
                cb.checked = false;
            });
            loadFromAppSheet(false);
        });

        const nhomList = document.getElementById("tbx-filter-nhom-thiet-bi-list");
        if (nhomList) nhomList.addEventListener("change", () => loadFromAppSheet(false));
        document.getElementById("tbx-nhom-check-all")?.addEventListener("click", () => {
            document.querySelectorAll("#tbx-filter-nhom-thiet-bi-list input[type='checkbox']").forEach((cb) => {
                cb.checked = true;
            });
            loadFromAppSheet(false);
        });
        document.getElementById("tbx-nhom-check-none")?.addEventListener("click", () => {
            document.querySelectorAll("#tbx-filter-nhom-thiet-bi-list input[type='checkbox']").forEach((cb) => {
                cb.checked = false;
            });
            loadFromAppSheet(false);
        });

        const search = document.getElementById("tbx-search");
        if (search) {
            search.addEventListener("input", () => {
                const q = search.value.trim().toLowerCase();
                const tbody = document.getElementById("tbx-tbody");
                if (!tbody) return;
                for (const tr of tbody.querySelectorAll("tr")) {
                    tr.style.display = !q || tr.textContent.toLowerCase().includes(q) ? "" : "none";
                }
            });
        }

        syncDateRangeBanner();
        loadFromAppSheet(false);
    }

    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
    else init();
})();
