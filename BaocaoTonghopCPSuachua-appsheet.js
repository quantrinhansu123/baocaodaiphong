/**
 * Báo cáo tổng hợp chi phí sửa chữa — lấy dữ liệu từ «Nhập xuất luân chuyển CT».
 * Nguồn cột: Tên link kiện sửa chữa, Ngày sửa chữa, ĐVT SC, Số lượng SC, Đơn giá SC, Thành tiền SC.
 */
(function () {
    "use strict";

    const APPSHEET_CONFIG = {
        appId: "3be5baea-960f-4d3f-b388-d13364cc4f22",
        accessKey: "V2-GaoRd-ItaM1-r44oH-c6Smd-uOe7V-cmVoK-IJINF-5XLQa"
    };

    const TABLE_NXLC_CT = "Nhập xuất luân chuyển CT";
    const SC_PREVIEW_ROW_CAP = 250;

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

    function escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    function normalizeSpaces(s) {
        return String(s ?? "").trim().replace(/\s+/g, " ");
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

    function formatMoney(n) {
        if (n == null || isNaN(n)) return "0";
        return new Intl.NumberFormat("vi-VN", { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(n);
    }

    function formatPercent(v) {
        if (v == null || !isFinite(v)) return "0.00%";
        return `${v.toFixed(2)}%`;
    }

    function parseDateFlexible(value) {
        const s = normalizeSpaces(cellDisplayString(value));
        if (!s) return null;
        let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
        if (m) {
            const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
            return isNaN(d.getTime()) ? null : d;
        }
        m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[\s,].*)?$/);
        if (m) {
            const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
            return isNaN(d.getTime()) ? null : d;
        }
        const native = new Date(s);
        if (!isNaN(native.getTime())) return native;
        return null;
    }

    function parseDateSlashSmart(s) {
        const m = String(s ?? "").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[\s,].*)?$/);
        if (!m) return null;
        const a = Number(m[1]);
        const b = Number(m[2]);
        const y = Number(m[3]);

        // Không mơ hồ: 25/4 => dd/mm, 4/25 => mm/dd.
        if (a > 12 && b <= 12) return new Date(y, b - 1, a);
        if (b > 12 && a <= 12) return new Date(y, a - 1, b);

        // Mơ hồ (cả 2 <=12): ưu tiên MM/DD theo dữ liệu AppSheet đang trả.
        return new Date(y, a - 1, b);
    }

    function parseNgaySuaChuaValue(raw) {
        if (raw == null) return null;

        // Ưu tiên chuỗi ISO trong object AppSheet để tránh đảo ngày/tháng.
        if (typeof raw === "object" && raw !== null) {
            const ordered = [
                raw.Value,
                raw.value,
                raw.DisplayValue,
                raw.displayValue,
                raw.FormattedValue,
                raw.formattedValue,
                raw.Text,
                raw.text
            ];
            for (const item of ordered) {
                const txt = normalizeSpaces(item);
                if (!txt) continue;
                const iso = txt.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
                if (iso) {
                    const dIso = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
                    if (!isNaN(dIso.getTime())) return dIso;
                }
                const dSlash = parseDateSlashSmart(txt);
                if (dSlash && !isNaN(dSlash.getTime())) return dSlash;
                const dFallback = parseDateFlexible(txt);
                if (dFallback && !isNaN(dFallback.getTime())) return dFallback;
            }
            return null;
        }

        const rawText = normalizeSpaces(raw);
        const dSlash = parseDateSlashSmart(rawText);
        if (dSlash && !isNaN(dSlash.getTime())) return dSlash;
        return parseDateFlexible(rawText);
    }

    function dateToDdMmYyyy(d) {
        if (!d || isNaN(d.getTime())) return "";
        return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
    }

    function normalizeScDateCellValue(v) {
        const d = parseNgaySuaChuaValue(v);
        return d ? dateToDdMmYyyy(d) : cellDisplayString(v);
    }

    function normalizeNxlcScDateFields(row) {
        if (!row || typeof row !== "object") return row;
        const out = { ...row };
        const dateKeys = [
            "Ngày sửa chữa",
            "Ngay sua chua",
            "Ngày SC",
            "Ngay SC"
        ];
        for (const key of dateKeys) {
            if (out[key] == null) continue;
            out[key] = normalizeScDateCellValue(out[key]);
        }
        return out;
    }

    function pickFirstCell(row, keys) {
        for (const k of keys) {
            const v = row?.[k];
            if (v == null) continue;
            const s = normalizeSpaces(cellDisplayString(v));
            if (s) return s;
        }
        return "";
    }

    function pickByRegex(row, pattern) {
        for (const [k, v] of Object.entries(row || {})) {
            if (!pattern.test(normalizeSpaces(k))) continue;
            const s = normalizeSpaces(cellDisplayString(v));
            if (s) return s;
        }
        return "";
    }

    function normalizeHeaderKey(s) {
        return String(s ?? "")
            .trim()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase()
            .replace(/\s+/g, " ");
    }

    function pickByHeaderAliases(row, aliases) {
        const want = new Set((aliases || []).map((a) => normalizeHeaderKey(a)));
        for (const [k, v] of Object.entries(row || {})) {
            if (!want.has(normalizeHeaderKey(k))) continue;
            const s = normalizeSpaces(cellDisplayString(v));
            if (s) return s;
        }
        return "";
    }

    function pickLoaiPhieu(row) {
        const direct = pickFirstCell(row, ["Loại phiếu", "Loai phieu", "Loại", "Loai"]);
        if (direct) return direct;
        return pickByRegex(row, /lo[aạ]i.*phi[eế]u/i);
    }

    function isXuatLoaiPhieu(row) {
        return normalizeHeaderKey(pickLoaiPhieu(row)) === "xuat";
    }

    function pickTenLinhKienSuaChua(row) {
        const direct = pickFirstCell(row, [
            "Tên link kiện sửa chữa",
            "Tên linh kiện sửa chữa",
            "Ten link kien sua chua",
            "Ten linh kien sua chua",
            "Tên LK sửa chữa",
            "Ten LK sua chua"
        ]);
        if (direct) return direct;
        return pickByRegex(row, /t[eê]n.*(linh|link).*(ki[eệ]n).*(s[ửu]a).*(ch[uữ]a)/i);
    }

    function pickNgaySuaChua(row) {
        const keys = [
            "Ngày sửa chữa",
            "Ngay sua chua",
            "Ngày SC",
            "Ngay SC"
        ];
        // Ưu tiên tuyệt đối cột «Ngày sửa chữa»/alias trực tiếp.
        for (const k of keys) {
            if (row?.[k] == null) continue;
            const d = parseNgaySuaChuaValue(row[k]);
            if (d) return d;
        }
        // Fallback khi nguồn đổi tên cột nhưng vẫn chứa đúng ngữ cảnh sửa chữa.
        for (const [k, v] of Object.entries(row || {})) {
            if (!/ng[aà]y.*(s[ửu]a|sc).*ch[uữ]a|ng[aà]y.*sc/i.test(normalizeSpaces(k))) continue;
            const d = parseNgaySuaChuaValue(v);
            if (d) return d;
        }
        return null;
    }

    function pickDvtSc(row) {
        const direct = pickFirstCell(row, ["ĐVT SC", "DVT SC", "Đơn vị SC", "Don vi SC"]);
        if (direct) return direct;
        return pickByRegex(row, /[đd]vt.*sc|đ[oơ]n.*v[iị].*sc/i);
    }

    function pickSoLuongSc(row) {
        const keys = ["Số lượng SC", "So luong SC", "SL SC", "So luong sua chua"];
        for (const k of keys) {
            const n = parseNumeric(row?.[k]);
            if (n != null) return n;
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (!/(s[oố].*l[ượ]ng.*sc|sl.*sc|s[oố].*l[ượ]ng.*s[ửu]a.*ch[uữ]a)/i.test(normalizeSpaces(k))) continue;
            const n = parseNumeric(v);
            if (n != null) return n;
        }
        return null;
    }

    function pickDonGiaSc(row) {
        const keys = ["Đơn giá SC", "Don gia SC", "Đơn giá sửa chữa", "Don gia sua chua"];
        for (const k of keys) {
            const n = parseNumeric(row?.[k]);
            if (n != null) return n;
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (!/(đ[oơ]n.*gi[aá].*sc|don.*gia.*sc|đ[oơ]n.*gi[aá].*s[ửu]a.*ch[uữ]a)/i.test(normalizeSpaces(k))) continue;
            const n = parseNumeric(v);
            if (n != null) return n;
        }
        return null;
    }

    function pickThanhTienSc(row) {
        const keys = ["Thành tiền SC", "Thanh tien SC", "Tiền SC", "Tien SC"];
        for (const k of keys) {
            const n = parseNumeric(row?.[k]);
            if (n != null) return n;
        }
        for (const [k, v] of Object.entries(row || {})) {
            if (!/(th[aà]nh.*ti[eề]n.*sc|ti[eề]n.*sc|th[aà]nh.*ti[eề]n.*s[ửu]a.*ch[uữ]a)/i.test(normalizeSpaces(k))) continue;
            const n = parseNumeric(v);
            if (n != null) return n;
        }
        return null;
    }

    function pickTenThietBiSuaChuaAgg(row) {
        const direct = pickFirstCell(row, [
            "Tên XMTB",
            "Ten XMTB",
            "Tên thiết bị sửa chữa",
            "Ten thiet bi sua chua"
        ]);
        if (direct) return direct;
        return pickFirstCell(row, ["Tên tài sản", "Ten tai san", "Tên thiết bị", "Ten thiet bi"]);
    }

    function pickTenLaiMayAgg(row) {
        const direct = pickFirstCell(row, [
            "Tên LM",
            "Ten LM",
            "Tên lái máy",
            "Ten lai may",
            "Lái máy",
            "Lai may",
            "Tên lái máy/XMTB",
            "Ten lai may/XMTB",
            "Tên người vận hành",
            "Ten nguoi van hanh"
        ]);
        if (direct) return direct;
        return pickByRegex(row, /(t[eê]n.*(l[aá]i|lm).*(m[aá]y|xmtb)|l[aá]i.*m[aá]y|t[eê]n.*v[aậ]n.*h[aà]nh)/i);
    }

    function pickNhomTextAgg(row) {
        const aliases = [
            "Tên nhóm",
            "Ten nhom",
            "TenNhom",
            "Tên loại",
            "Ten loai",
            "Loại",
            "Loai",
            "Nhóm thiết bị",
            "Nhom thiet bi",
            "Nhóm xe máy thiết bị",
            "Nhom xe may thiet bi",
            "Nhóm XMTB",
            "Nhom XMTB",
            "Loại máy",
            "Loai may",
            "Tên loại xe",
            "Ten loai xe"
        ];
        const nhom = pickFirstCell(row, aliases);
        if (nhom) return nhom;
        const nhomAlias = pickByHeaderAliases(row, aliases);
        if (nhomAlias) return nhomAlias;
        const nhomRegex = pickByRegex(row, /t[eê]n.*nh[oô]m|t[eê]n.*lo[aạ]i|lo[aạ]i.*(m[aá]y|xe|thi[eế]t\s*b[iị]|xmtb)|nh[oô]m.*(th[iệ]t\s*b[iị]|xe\s*m[aá]y|xmtb)/i);
        if (nhomRegex) return nhomRegex;
        return pickFirstCell(row, ["Tên thiết bị", "Ten thiet bi", "TenThietBi"]);
    }

    function getNxlcTextFilterNeedles() {
        const g = (id) => normalizeSpaces(document.getElementById(id)?.value ?? "").toLowerCase();
        return {
            tenTb: g("sc-filter-ten-thiet-bi"),
            tenLm: g("sc-filter-ten-lai-may"),
            nhom: g("sc-filter-nhom")
        };
    }

    function rowMatchesNxlcTextFilters(row, needles) {
        const n = needles || getNxlcTextFilterNeedles();
        const hayTb = normalizeSpaces(pickTenThietBiSuaChuaAgg(row)).toLowerCase();
        const hayLm = normalizeSpaces(pickTenLaiMayAgg(row)).toLowerCase();
        const hayNhom = normalizeSpaces(pickNhomTextAgg(row)).toLowerCase();
        if (n.tenTb && hayTb !== n.tenTb) return false;
        if (n.tenLm && hayLm !== n.tenLm) return false;
        if (n.nhom && hayNhom !== n.nhom) return false;
        return true;
    }

    function fillRepairFilterSelect(el, values, emptyLabel) {
        if (!el || el.tagName !== "SELECT") return;
        const prev = el.value;
        el.innerHTML = "";
        const o0 = document.createElement("option");
        o0.value = "";
        o0.textContent = emptyLabel || "— Tất cả —";
        el.appendChild(o0);
        for (const v of values) {
            const t = normalizeSpaces(v);
            if (!t) continue;
            const o = document.createElement("option");
            o.value = t;
            o.textContent = t;
            el.appendChild(o);
        }
        const hasPrev = prev && [...el.options].some((op) => op.value === prev);
        el.value = hasPrev ? prev : "";
    }

    function populateNxlcFilterSelects(rows) {
        const setTb = new Set();
        const setLm = new Set();
        const setNhom = new Set();
        for (const row of rows || []) {
            const tb = normalizeSpaces(pickTenThietBiSuaChuaAgg(row));
            const lm = normalizeSpaces(pickTenLaiMayAgg(row));
            const nh = normalizeSpaces(pickNhomTextAgg(row));
            if (tb) setTb.add(tb);
            if (lm) setLm.add(lm);
            if (nh) setNhom.add(nh);
        }
        const sortVi = (a, b) => String(a).localeCompare(String(b), "vi");
        fillRepairFilterSelect(document.getElementById("sc-filter-ten-thiet-bi"), [...setTb].sort(sortVi));
        fillRepairFilterSelect(document.getElementById("sc-filter-ten-lai-may"), [...setLm].sort(sortVi));
        fillRepairFilterSelect(document.getElementById("sc-filter-nhom"), [...setNhom].sort(sortVi));
    }

    function parseYmdToLocalStart(ymd) {
        const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ymd?.trim() ?? "");
        if (!m) return null;
        const y = +m[1], mo = +m[2] - 1, d = +m[3];
        const t = new Date(y, mo, d, 0, 0, 0, 0);
        return isNaN(t.getTime()) ? null : t;
    }

    function parseYmdToLocalEnd(ymd) {
        const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ymd?.trim() ?? "");
        if (!m) return null;
        const y = +m[1], mo = +m[2] - 1, d = +m[3];
        const t = new Date(y, mo, d, 23, 59, 59, 999);
        return isNaN(t.getTime()) ? null : t;
    }

    function getDateRangeFromFilters() {
        let fromRaw = document.getElementById("sc-filter-from-date")?.value?.trim() ?? "";
        let toRaw = document.getElementById("sc-filter-to-date")?.value?.trim() ?? "";
        if (fromRaw && toRaw) {
            const a = parseYmdToLocalStart(fromRaw);
            const b = parseYmdToLocalStart(toRaw);
            if (a && b && a.getTime() > b.getTime()) {
                [fromRaw, toRaw] = [toRaw, fromRaw];
                const fromEl = document.getElementById("sc-filter-from-date");
                const toEl = document.getElementById("sc-filter-to-date");
                if (fromEl) fromEl.value = fromRaw;
                if (toEl) toEl.value = toRaw;
            }
        }
        const from = fromRaw ? parseYmdToLocalStart(fromRaw) : null;
        const to = toRaw ? parseYmdToLocalEnd(toRaw) : null;
        return {
            hasFrom: !!from,
            hasTo: !!to,
            from,
            to
        };
    }

    function rowInDateRange(dateObj, dateRange) {
        if (!dateRange?.hasFrom && !dateRange?.hasTo) return true;
        if (!dateObj || isNaN(dateObj.getTime())) return false;
        if (dateRange.hasFrom && dateObj < dateRange.from) return false;
        if (dateRange.hasTo && dateObj > dateRange.to) return false;
        return true;
    }

    function formatDateDdMmYyyy(d) {
        return dateToDdMmYyyy(d);
    }

    function syncDateRangeBanner() {
        const el = document.getElementById("sc-date-range");
        if (!el) return;
        const fromRaw = document.getElementById("sc-filter-from-date")?.value?.trim() ?? "";
        const toRaw = document.getElementById("sc-filter-to-date")?.value?.trim() ?? "";
        const dFrom = fromRaw ? parseYmdToLocalStart(fromRaw) : null;
        const dTo = toRaw ? parseYmdToLocalStart(toRaw) : null;
        const labelFrom = dFrom ? formatDateDdMmYyyy(dFrom) : "…";
        const labelTo = dTo ? formatDateDdMmYyyy(dTo) : "…";
        el.textContent = `(Từ ngày ${labelFrom} tới ngày ${labelTo})`;
    }

    function aggregateRepairRows(rows) {
        const dateRange = getDateRangeFromFilters();

        const filteredRows = [];
        const textNeedles = getNxlcTextFilterNeedles();
        for (const row of rows || []) {
            const ten = pickTenLinhKienSuaChua(row);
            const ngay = pickNgaySuaChua(row);
            if (!ten || !rowInDateRange(ngay, dateRange)) continue;
            if (!rowMatchesNxlcTextFilters(row, textNeedles)) continue;
            filteredRows.push({ raw: row, ten, ngay });
        }

        let years = [];
        if (dateRange.hasFrom && dateRange.hasTo) {
            for (let y = dateRange.from.getFullYear(); y <= dateRange.to.getFullYear(); y += 1) years.push(y);
        } else {
            const ys = new Set(filteredRows.map((r) => r.ngay.getFullYear()));
            years = [...ys].sort((a, b) => a - b);
            if (!years.length) {
                if (dateRange.hasFrom) years = [dateRange.from.getFullYear()];
                else if (dateRange.hasTo) years = [dateRange.to.getFullYear()];
            }
        }
        if (!years.length) years = [new Date().getFullYear()];

        const yearSet = new Set(years);
        const map = new Map();
        const totalsByYear = new Map(years.map((y) => [y, 0]));

        for (const item of filteredRows) {
            const year = item.ngay.getFullYear();
            if (!yearSet.has(year)) continue;
            const row = item.raw;
            const dvt = pickDvtSc(row);
            const qty = pickSoLuongSc(row);
            const price = pickDonGiaSc(row);
            const thanhTienRaw = pickThanhTienSc(row);
            const thanhTien = thanhTienRaw != null ? thanhTienRaw : (qty != null && price != null ? qty * price : 0);
            if (!isFinite(thanhTien) || thanhTien === 0) continue;

            const key = item.ten.toLowerCase();
            const prev = map.get(key) || {
                ten: item.ten,
                dvt,
                byYear: {}
            };
            if (!prev.dvt && dvt) prev.dvt = dvt;
            prev.byYear[year] = (prev.byYear[year] || 0) + thanhTien;
            totalsByYear.set(year, (totalsByYear.get(year) || 0) + thanhTien);
            map.set(key, prev);
        }

        const rowsOut = [...map.values()]
            .map((r) => ({
                ...r,
                total: years.reduce((sum, y) => sum + (r.byYear[y] || 0), 0)
            }))
            .filter((r) => r.total > 0)
            .sort((a, b) => b.total - a.total || a.ten.localeCompare(b.ten, "vi"));

        const grandTotal = years.reduce((sum, y) => sum + (totalsByYear.get(y) || 0), 0);
        return { rows: rowsOut, years, totalsByYear, grandTotal, dateRange };
    }

    function renderRepairSummaryTable() {
        const thead = document.getElementById("sc-thead") || document.querySelector(".repair-summary-table thead");
        const tbody = document.getElementById("sc-tbody") || document.querySelector(".repair-summary-table tbody");
        if (!thead || !tbody) return;

        const agg = aggregateRepairRows(cacheNxlcCtRows || []);
        const yearHeaders = agg.years || [];
        const colCount = 1 + yearHeaders.length * 2 + 1;

        const h1 = `<tr>
            <th rowspan="2" style="width: 34%;">DIỄN GIẢI</th>
            ${yearHeaders.map((y) => `<th colspan="2" style="width: ${Math.max(12, Math.floor(52 / Math.max(1, yearHeaders.length)))}%;">Năm ${y}</th>`).join("")}
            <th rowspan="2" style="width: 14%;">Total Số tiền SC</th>
        </tr>`;
        const h2 = `<tr class="sub-head">${yearHeaders.map(() => "<th>Số tiền SC</th><th>Tỷ trọng</th>").join("")}</tr>`;
        thead.innerHTML = h1 + h2;

        if (!agg.rows.length) {
            tbody.innerHTML = `<tr class="child-row"><td colspan="${colCount}" class="text-left">Không có dữ liệu phù hợp bộ lọc từ «Nhập xuất luân chuyển CT».</td></tr>`;
            return;
        }

        const pct = (v, y) => {
            const total = agg.totalsByYear.get(y) || 0;
            return total > 0 ? (v / total) * 100 : 0;
        };

        const rowsHtml = agg.rows
            .map((r) => {
                const nameWithUnit = r.dvt ? `${r.ten} (${r.dvt})` : r.ten;
                const yearCells = yearHeaders
                    .map((y) => {
                        const val = r.byYear[y] || 0;
                        return `<td class="text-right">${escapeHtml(formatMoney(val))}</td><td>${escapeHtml(formatPercent(pct(val, y)))}</td>`;
                    })
                    .join("");
                return `<tr class="child-row">
                    <td class="text-left">${escapeHtml(nameWithUnit)}</td>
                    ${yearCells}
                    <td class="text-right">${escapeHtml(formatMoney(r.total))}</td>
                </tr>`;
            })
            .join("");

        const totalCells = yearHeaders
            .map((y) => {
                const val = agg.totalsByYear.get(y) || 0;
                return `<td class="text-right">${escapeHtml(formatMoney(val))}</td><td>${escapeHtml(formatPercent(val > 0 ? 100 : 0))}</td>`;
            })
            .join("");
        const totalRow = `<tr class="grand-total-row">
            <td class="text-left">TỔNG CỘNG</td>
            ${totalCells}
            <td class="text-right">${escapeHtml(formatMoney(agg.grandTotal))}</td>
        </tr>`;

        tbody.innerHTML = rowsHtml + totalRow;
    }

    function collectKeysFromRows(rows) {
        const set = new Set();
        for (const row of rows || []) {
            for (const key of Object.keys(row || {})) set.add(key);
        }
        const keys = [...set];
        const dateKeys = keys
            .filter((k) => /ng[aà]y.*(s[ửu]a|sc).*ch[uữ]a|ng[aà]y.*sc/i.test(String(k)))
            .sort((a, b) => String(a).localeCompare(String(b), "vi"));
        const rest = keys
            .filter((k) => !/ng[aà]y.*(s[ửu]a|sc).*ch[uữ]a|ng[aà]y.*sc/i.test(String(k)))
            .sort((a, b) => String(a).localeCompare(String(b), "vi"));
        return [...dateKeys, ...rest];
    }

    function buildNxlcPreviewTableHtml(rows) {
        const all = rows ?? [];
        if (!all.length) return '<p class="sc-prev-note">Không có dòng.</p>';

        const dropDateAliases = (k) => /ng[aà]y.*(s[ửu]a|sc).*ch[uữ]a|ng[aà]y.*sc/i.test(String(k));
        const rawKeys = collectKeysFromRows(all).filter((k) => !dropDateAliases(k));
        const keys = ["Ngày sửa chữa", ...rawKeys];

        const show = all.slice(0, SC_PREVIEW_ROW_CAP);
        const note =
            (all.length > SC_PREVIEW_ROW_CAP
                ? `Hiển thị tối đa ${SC_PREVIEW_ROW_CAP}/${all.length} dòng. `
                : `${all.length} dòng. `) + "Cột ngày sửa chữa đã ép chuẩn dd/mm/yyyy.";

        const head = `<tr>${keys.map((k) => `<th>${escapeHtml(k)}</th>`).join("")}</tr>`;
        const body = show
            .map((row) => {
                const d = pickNgaySuaChua(row);
                const dateCell = `<td>${escapeHtml(d ? dateToDdMmYyyy(d) : "")}</td>`;
                const restCells = rawKeys.map((k) => `<td>${escapeHtml(cellDisplayString(row[k]))}</td>`).join("");
                return `<tr>${dateCell}${restCells}</tr>`;
            })
            .join("");
        return `<p class="sc-prev-note">${escapeHtml(note)}</p><div class="sc-prev-scroll"><table class="sc-prev-table"><thead>${head}</thead><tbody>${body}</tbody></table></div>`;
    }

    function buildRawPreviewTableHtml(rows) {
        const all = rows ?? [];
        if (!all.length) return '<p class="sc-prev-note">Không có dòng.</p>';
        const keys = collectKeysFromRows(all);
        if (!keys.includes("Ngày sửa chữa")) keys.unshift("Ngày sửa chữa");
        if (!keys.length) return `<p class="sc-prev-note">${all.length} dòng nhưng không có cột hiển thị.</p>`;

        const show = all.slice(0, SC_PREVIEW_ROW_CAP);
        const note =
            (all.length > SC_PREVIEW_ROW_CAP
                ? `Hiển thị tối đa ${SC_PREVIEW_ROW_CAP}/${all.length} dòng. `
                : `${all.length} dòng. `) + "Kéo ngang để xem đầy đủ cột.";

        const head = `<tr>${keys.map((k) => `<th>${escapeHtml(k)}</th>`).join("")}</tr>`;
        const body = show
            .map((row) => {
                const cells = keys
                    .map((k) => {
                        if (k === "Ngày sửa chữa") {
                            const direct = normalizeSpaces(cellDisplayString(row[k]));
                            const d = pickNgaySuaChua(row);
                            return `<td>${escapeHtml(d ? dateToDdMmYyyy(d) : direct)}</td>`;
                        }
                        return `<td>${escapeHtml(cellDisplayString(row[k]))}</td>`;
                    })
                    .join("");
                return `<tr>${cells}</tr>`;
            })
            .join("");
        return `<p class="sc-prev-note">${escapeHtml(note)}</p><div class="sc-prev-scroll"><table class="sc-prev-table"><thead>${head}</thead><tbody>${body}</tbody></table></div>`;
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
        if (Array.isArray(data)) {
            const rows = data.map(normalizeNxlcScDateFields);
            return tableName === TABLE_NXLC_CT ? rows.filter(isXuatLoaiPhieu) : rows;
        }
        if (data && Array.isArray(data.Rows)) {
            const rows = data.Rows.map(normalizeNxlcScDateFields);
            return tableName === TABLE_NXLC_CT ? rows.filter(isXuatLoaiPhieu) : rows;
        }
        return [];
    }

    function setStatus(msg, isErr) {
        const el = document.getElementById("sc-appsheet-status");
        if (!el) return;
        el.textContent = msg;
        el.classList.toggle("err", !!isErr);
    }

    async function loadFromAppSheet(forceRefresh) {
        const btn = document.getElementById("sc-btn-load");
        const refBtn = document.getElementById("sc-btn-refresh");
        if (btn) btn.disabled = true;
        if (refBtn) refBtn.disabled = true;

        try {
            if (forceRefresh) cacheNxlcCtRows = null;
            setStatus("Đang tải AppSheet…", false);
            if (!cacheNxlcCtRows) {
                setStatus(`Đang tải «${TABLE_NXLC_CT}»…`, false);
                cacheNxlcCtRows = await fetchAppSheetTable(TABLE_NXLC_CT);
            }
            populateNxlcFilterSelects(cacheNxlcCtRows || []);
            renderRepairSummaryTable();
            setStatus(`Đã tải ${(cacheNxlcCtRows || []).length} dòng «${TABLE_NXLC_CT}».`, false);
        } catch (error) {
            console.error(error);
            setStatus(`Lỗi: ${error?.message || error}`, true);
            const tbody = document.getElementById("sc-tbody") || document.querySelector(".repair-summary-table tbody");
            const colCount = (document.querySelectorAll("#sc-thead th").length || 6);
            if (tbody) tbody.innerHTML = `<tr class="child-row"><td colspan="${colCount}" class="text-left">Không tải được dữ liệu AppSheet.</td></tr>`;
        } finally {
            if (btn) btn.disabled = false;
            if (refBtn) refBtn.disabled = false;
        }
    }

    function openLoadedDataModal() {
        const modal = document.getElementById("sc-loaded-data-modal");
        if (!modal) return;
        if (cacheNxlcCtRows == null) {
            alert("Chưa tải dữ liệu AppSheet. Bấm «Tải dữ liệu AppSheet» trước.");
            return;
        }
        const agg = aggregateRepairRows(cacheNxlcCtRows || []);
        const yearsLabel = (agg.years || []).join(", ");
        const dateText =
            agg.dateRange?.hasFrom || agg.dateRange?.hasTo
                ? `${agg.dateRange?.hasFrom ? formatDateDdMmYyyy(agg.dateRange.from) : "…"} -> ${agg.dateRange?.hasTo ? formatDateDdMmYyyy(agg.dateRange.to) : "…"}`
                : "không giới hạn";
        const previewRows = (cacheNxlcCtRows || []).filter((r) => rowInDateRange(pickNgaySuaChua(r), agg.dateRange));
        const panel = document.getElementById("sc-loaded-panel-nxlc");
        if (panel) {
            panel.innerHTML =
                `<p class="sc-prev-note">Khoảng lọc: ${escapeHtml(dateText)} | Năm hiển thị: ${escapeHtml(yearsLabel || "—")}.</p>` +
                buildNxlcPreviewTableHtml(previewRows);
        }
        const tab = document.getElementById("sc-tab-nxlc");
        if (tab) tab.textContent = `Nhập xuất luân chuyển CT (${previewRows.length}/${(cacheNxlcCtRows || []).length})`;
        modal.style.display = "flex";
        modal.setAttribute("aria-hidden", "false");
        document.body.style.overflow = "hidden";
    }

    function closeLoadedDataModal() {
        const modal = document.getElementById("sc-loaded-data-modal");
        if (!modal) return;
        modal.style.display = "none";
        modal.setAttribute("aria-hidden", "true");
        document.body.style.overflow = "";
    }

    function init() {
        document.getElementById("sc-btn-load")?.addEventListener("click", () => loadFromAppSheet(false));
        document.getElementById("sc-btn-refresh")?.addEventListener("click", () => loadFromAppSheet(true));
        document.getElementById("sc-btn-view-loaded-tables")?.addEventListener("click", openLoadedDataModal);
        document.getElementById("sc-loaded-close")?.addEventListener("click", closeLoadedDataModal);
        document.getElementById("sc-loaded-data-modal")?.addEventListener("click", (e) => {
            if (e.target && e.target.id === "sc-loaded-data-modal") closeLoadedDataModal();
        });
        document.addEventListener("keydown", (e) => {
            if (e.key !== "Escape") return;
            const modal = document.getElementById("sc-loaded-data-modal");
            if (modal && modal.style.display === "flex") closeLoadedDataModal();
        });

        const rerender = () => {
            syncDateRangeBanner();
            renderRepairSummaryTable();
        };
        document.getElementById("sc-filter-from-date")?.addEventListener("change", rerender);
        document.getElementById("sc-filter-to-date")?.addEventListener("change", rerender);
        const textRerender = () => {
            syncDateRangeBanner();
            renderRepairSummaryTable();
        };
        for (const id of ["sc-filter-ten-thiet-bi", "sc-filter-ten-lai-may", "sc-filter-nhom"]) {
            document.getElementById(id)?.addEventListener("change", textRerender);
        }
        syncDateRangeBanner();

        loadFromAppSheet(false);
    }

    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
    else init();
})();
