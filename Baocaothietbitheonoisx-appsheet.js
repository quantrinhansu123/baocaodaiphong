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

    let cacheDsRows = null;

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

    function rowMatchesFilters(row, filters) {
        const parts = [
            filters.congTrinh,
            filters.nhomNcc,
            filters.tenNcc,
            filters.tenKho,
            filters.nhomVt,
            filters.tenVt
        ].filter((x) => x && String(x).trim());
        if (!parts.length) return true;
        let blob;
        try {
            blob = JSON.stringify(row).toLowerCase();
        } catch (e) {
            return true;
        }
        return parts.every((p) => blob.includes(String(p).trim().toLowerCase()));
    }

    function getFilters() {
        return {
            congTrinh: document.getElementById("tbx-filter-cong-trinh")?.value?.trim() ?? "",
            nhomNcc: document.getElementById("tbx-filter-nhom-ncc")?.value?.trim() ?? "",
            tenNcc: document.getElementById("tbx-filter-ten-ncc")?.value?.trim() ?? "",
            tenKho: document.getElementById("tbx-filter-kho")?.value?.trim() ?? "",
            nhomVt: document.getElementById("tbx-filter-nhom-vt")?.value?.trim() ?? "",
            tenVt: document.getElementById("tbx-filter-ten-vt")?.value?.trim() ?? ""
        };
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

    function renderTable(groups) {
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
                addTot(st);
                const ma = pickDsMaThietBi(r);
                const ten = pickDsTenThietBi(r);
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
            if (forceRefresh) cacheDsRows = null;
            if (!cacheDsRows) {
                setStatus(`Đang tải «${TABLE_DS_TAI_SAN}»…`, false);
                cacheDsRows = await fetchAppSheetTable(TABLE_DS_TAI_SAN);
            }
            const filters = getFilters();
            const groups = buildGroupedReport(cacheDsRows, filters);
            renderTable(groups);
            syncDateRangeBanner();
            const n = cacheDsRows?.length ?? 0;
            const g = groups.length;
            setStatus(`Đã tải ${n} dòng «${TABLE_DS_TAI_SAN}» — ${g} nhóm nơi quản lý.`, false);
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

        for (const id of ["tbx-filter-from-date", "tbx-filter-to-date"]) {
            const el = document.getElementById(id);
            if (!el) continue;
            const fn = () => syncDateRangeBanner();
            el.addEventListener("input", fn);
            el.addEventListener("change", fn);
            el.addEventListener("blur", fn);
        }

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
