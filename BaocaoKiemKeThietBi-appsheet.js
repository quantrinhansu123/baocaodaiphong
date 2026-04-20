(function () {
    "use strict";

    const APPSHEET_CONFIG = {
        appId: "3be5baea-960f-4d3f-b388-d13364cc4f22",
        accessKey: "V2-GaoRd-ItaM1-r44oH-c6Smd-uOe7V-cmVoK-IJINF-5XLQa"
    };

    const TABLE_NXLC_CT = "Nhập xuất luân chuyển CT";
    const PREVIEW_ROW_CAP = 500;
    const COL_COUNT = 13;

    let cacheRows = null;

    function setStatus(message, isError) {
        const el = document.getElementById("kktb-appsheet-status");
        if (!el) return;
        el.textContent = message || "";
        el.classList.toggle("text-red-600", !!isError);
        el.classList.toggle("text-[#2b7132]", !isError);
    }

    function cellString(value) {
        if (value == null) return "";
        if (typeof value === "object" && value !== null && "DisplayValue" in value) return String(value.DisplayValue ?? "");
        if (typeof value === "object" && value !== null && "displayValue" in value) return String(value.displayValue ?? "");
        if (typeof value === "object" && value !== null && "Value" in value) return String(value.Value ?? "");
        if (typeof value === "object" && value !== null && "value" in value) return String(value.value ?? "");
        return String(value);
    }

    function parseNumeric(raw) {
        if (raw == null || raw === "") return null;
        if (typeof raw === "number" && !isNaN(raw)) return raw;
        const source = cellString(raw).trim();
        if (!source) return null;
        const match = source.match(/-?[\d.,]+/);
        if (!match) return null;
        let normalized = match[0];
        const comma = normalized.indexOf(",");
        const dot = normalized.indexOf(".");
        if (comma !== -1 && dot !== -1) {
            if (comma > dot) normalized = normalized.replace(/\./g, "").replace(",", ".");
            else normalized = normalized.replace(/,/g, "");
        } else if (comma !== -1) {
            normalized = normalized.replace(",", ".");
        } else if (dot !== -1 && normalized.indexOf(".", dot + 1) !== -1) {
            normalized = normalized.replace(/\./g, "");
        }
        const number = parseFloat(normalized);
        return isNaN(number) ? null : number;
    }

    function formatNum(value) {
        if (value == null || isNaN(value)) return "";
        return new Intl.NumberFormat("vi-VN", { maximumFractionDigits: 2 }).format(value);
    }

    function escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    function safeKey(value) {
        return String(value ?? "")
            .trim()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase()
            .replace(/[^a-z0-9]/g, "");
    }

    function pickFirstCell(row, keys) {
        for (const key of keys) {
            const value = row?.[key];
            if (value == null || value === "") continue;
            const text = cellString(value).trim();
            if (text) return text;
        }
        return "";
    }

    function pickNumericByAliases(row, exactKeys, regexFallback) {
        for (const key of exactKeys) {
            const value = parseNumeric(row?.[key]);
            if (value != null) return value;
        }
        if (regexFallback) {
            for (const [key, value] of Object.entries(row || {})) {
                if (!regexFallback.test(String(key))) continue;
                const parsed = parseNumeric(value);
                if (parsed != null) return parsed;
            }
        }
        return null;
    }

    function extractRecord(row) {
        const ten = pickFirstCell(row, [
            "Tên XMTB",
            "Tên thiết bị",
            "Ten thiet bi",
            "Tên tài sản",
            "Ten tai san",
            "Tên NL",
            "Ten NL"
        ]);
        const ma = pickFirstCell(row, ["Mã số", "Ma so", "Mã thiết bị", "Ma thiet bi", "Mã tài sản", "Ma tai san", "Mã", "Ma"]);
        const nhom = pickFirstCell(row, ["Nhóm XMTB", "Nhom XMTB", "Nhóm thiết bị", "Nhom thiet bi", "Nhóm", "Nhom"]);
        const ghiChu = pickFirstCell(row, ["Ghi chú", "Ghi chu", "GhiChu"]);

        const theoSoSl = pickNumericByAliases(
            row,
            ["Theo sổ SL", "Theo so SL", "SL theo sổ", "SL theo so", "Sl tồn ĐK", "SL tồn ĐK", "SL tồn DK", "SL"],
            /theo.*so.*sl|sl.*theo.*so|ton.*d[au]k|sl\s*ton/i
        );
        const theoSoNguyenGia = pickNumericByAliases(
            row,
            ["Theo sổ Nguyên giá", "Theo so Nguyen gia", "Nguyên giá theo sổ", "Nguyen gia theo so", "Nguyên giá", "Nguyen gia", "Đơn giá", "Don gia"],
            /theo.*so.*nguyen.*gia|nguyen.*gia|don.*gia/
        );
        const theoSoGiaTri = pickNumericByAliases(
            row,
            ["Theo sổ Giá trị", "Theo so Gia tri", "Giá trị theo sổ", "Gia tri theo so", "Giá trị", "Gia tri", "Thành tiền", "Thanh tien", "Thành tiền NL"],
            /theo.*so.*gia.*tri|gia.*tri.*theo.*so|thanh.*tien|gia.*tri/
        );

        const kiemKeSl = pickNumericByAliases(
            row,
            ["Kiểm kê SL", "Kiem ke SL", "SL kiểm kê", "SL kiem ke", "SL thực kê", "SL thuc ke"],
            /kiem.*ke.*sl|sl.*kiem.*ke|thuc.*ke.*sl/
        );
        const kiemKeNguyenGia = pickNumericByAliases(
            row,
            ["Kiểm kê Nguyên giá", "Kiem ke Nguyen gia", "Nguyên giá kiểm kê", "Nguyen gia kiem ke"],
            /kiem.*ke.*nguyen.*gia|nguyen.*gia.*kiem.*ke/
        );
        const kiemKeGiaTri = pickNumericByAliases(
            row,
            ["Kiểm kê Giá trị", "Kiem ke Gia tri", "Giá trị kiểm kê", "Gia tri kiem ke"],
            /kiem.*ke.*gia.*tri|gia.*tri.*kiem.*ke/
        );

        let chenhLechSl = pickNumericByAliases(
            row,
            ["Chênh lệch SL", "Chenh lech SL", "SL chênh lệch", "SL chenh lech"],
            /chenh.*lech.*sl|sl.*chenh.*lech/
        );
        let chenhLechNguyenGia = pickNumericByAliases(
            row,
            ["Chênh lệch Nguyên giá", "Chenh lech Nguyen gia", "Nguyên giá chênh lệch", "Nguyen gia chenh lech"],
            /chenh.*lech.*nguyen.*gia|nguyen.*gia.*chenh.*lech/
        );
        let chenhLechGiaTri = pickNumericByAliases(
            row,
            ["Chênh lệch Giá trị", "Chenh lech Gia tri", "Giá trị chênh lệch", "Gia tri chenh lech"],
            /chenh.*lech.*gia.*tri|gia.*tri.*chenh.*lech/
        );

        if (chenhLechSl == null && kiemKeSl != null && theoSoSl != null) chenhLechSl = kiemKeSl - theoSoSl;
        if (chenhLechNguyenGia == null && kiemKeNguyenGia != null && theoSoNguyenGia != null) chenhLechNguyenGia = kiemKeNguyenGia - theoSoNguyenGia;
        if (chenhLechGiaTri == null && kiemKeGiaTri != null && theoSoGiaTri != null) chenhLechGiaTri = kiemKeGiaTri - theoSoGiaTri;

        return {
            ten,
            ma,
            nhom: nhom || "CHƯA PHÂN NHÓM",
            ghiChu,
            theoSoSl,
            theoSoNguyenGia,
            theoSoGiaTri,
            kiemKeSl,
            kiemKeNguyenGia,
            kiemKeGiaTri,
            chenhLechSl,
            chenhLechNguyenGia,
            chenhLechGiaTri
        };
    }

    function addNumber(target, field, value) {
        if (value == null || isNaN(value)) return;
        target[field] += value;
    }

    function toNumberOrNull(value) {
        return value == null || isNaN(value) ? null : value;
    }

    function aggregateRows(rows) {
        const grouped = new Map();
        for (const raw of rows || []) {
            const row = extractRecord(raw);
            const uniqueKey = `${safeKey(row.nhom)}::${safeKey(row.ma)}::${safeKey(row.ten)}`;
            if (!row.ten && !row.ma) continue;
            if (!grouped.has(uniqueKey)) {
                grouped.set(uniqueKey, {
                    ...row,
                    theoSoSl: 0,
                    theoSoNguyenGia: 0,
                    theoSoGiaTri: 0,
                    kiemKeSl: 0,
                    kiemKeNguyenGia: 0,
                    kiemKeGiaTri: 0,
                    chenhLechSl: 0,
                    chenhLechNguyenGia: 0,
                    chenhLechGiaTri: 0,
                    hasTheoSoSl: false,
                    hasTheoSoNguyenGia: false,
                    hasTheoSoGiaTri: false,
                    hasKiemKeSl: false,
                    hasKiemKeNguyenGia: false,
                    hasKiemKeGiaTri: false,
                    hasChenhLechSl: false,
                    hasChenhLechNguyenGia: false,
                    hasChenhLechGiaTri: false
                });
            }
            const cur = grouped.get(uniqueKey);
            addNumber(cur, "theoSoSl", row.theoSoSl);
            addNumber(cur, "theoSoNguyenGia", row.theoSoNguyenGia);
            addNumber(cur, "theoSoGiaTri", row.theoSoGiaTri);
            addNumber(cur, "kiemKeSl", row.kiemKeSl);
            addNumber(cur, "kiemKeNguyenGia", row.kiemKeNguyenGia);
            addNumber(cur, "kiemKeGiaTri", row.kiemKeGiaTri);
            addNumber(cur, "chenhLechSl", row.chenhLechSl);
            addNumber(cur, "chenhLechNguyenGia", row.chenhLechNguyenGia);
            addNumber(cur, "chenhLechGiaTri", row.chenhLechGiaTri);

            if (row.theoSoSl != null) cur.hasTheoSoSl = true;
            if (row.theoSoNguyenGia != null) cur.hasTheoSoNguyenGia = true;
            if (row.theoSoGiaTri != null) cur.hasTheoSoGiaTri = true;
            if (row.kiemKeSl != null) cur.hasKiemKeSl = true;
            if (row.kiemKeNguyenGia != null) cur.hasKiemKeNguyenGia = true;
            if (row.kiemKeGiaTri != null) cur.hasKiemKeGiaTri = true;
            if (row.chenhLechSl != null) cur.hasChenhLechSl = true;
            if (row.chenhLechNguyenGia != null) cur.hasChenhLechNguyenGia = true;
            if (row.chenhLechGiaTri != null) cur.hasChenhLechGiaTri = true;

            if (!cur.ghiChu && row.ghiChu) cur.ghiChu = row.ghiChu;
        }

        const byNhom = new Map();
        for (const item of grouped.values()) {
            item.theoSoSl = item.hasTheoSoSl ? item.theoSoSl : null;
            item.theoSoNguyenGia = item.hasTheoSoNguyenGia ? item.theoSoNguyenGia : null;
            item.theoSoGiaTri = item.hasTheoSoGiaTri ? item.theoSoGiaTri : null;
            item.kiemKeSl = item.hasKiemKeSl ? item.kiemKeSl : null;
            item.kiemKeNguyenGia = item.hasKiemKeNguyenGia ? item.kiemKeNguyenGia : null;
            item.kiemKeGiaTri = item.hasKiemKeGiaTri ? item.kiemKeGiaTri : null;
            item.chenhLechSl = item.hasChenhLechSl ? item.chenhLechSl : null;
            item.chenhLechNguyenGia = item.hasChenhLechNguyenGia ? item.chenhLechNguyenGia : null;
            item.chenhLechGiaTri = item.hasChenhLechGiaTri ? item.chenhLechGiaTri : null;

            if (!byNhom.has(item.nhom)) byNhom.set(item.nhom, []);
            byNhom.get(item.nhom).push(item);
        }

        const sortedGroups = [...byNhom.entries()].sort((a, b) => a[0].localeCompare(b[0], "vi"));
        for (const [, items] of sortedGroups) {
            items.sort((a, b) => `${a.ten} ${a.ma}`.localeCompare(`${b.ten} ${b.ma}`, "vi"));
        }
        return sortedGroups;
    }

    function renderReport(rows) {
        const tbody = document.getElementById("kktb-tbody");
        if (!tbody) return;

        if (!Array.isArray(rows) || rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="${COL_COUNT}" class="text-center" style="padding:10px;">Không có dữ liệu từ AppSheet.</td></tr>`;
            return;
        }

        const groups = aggregateRows(rows);
        if (groups.length === 0) {
            tbody.innerHTML = `<tr><td colspan="${COL_COUNT}" class="text-center" style="padding:10px;">Không tìm thấy cột phù hợp để dựng báo cáo từ bảng "${escapeHtml(TABLE_NXLC_CT)}".</td></tr>`;
            return;
        }

        const totals = {
            theoSoSl: 0,
            theoSoNguyenGia: 0,
            theoSoGiaTri: 0,
            kiemKeSl: 0,
            kiemKeNguyenGia: 0,
            kiemKeGiaTri: 0,
            chenhLechSl: 0,
            chenhLechNguyenGia: 0,
            chenhLechGiaTri: 0
        };

        let html = "";
        groups.forEach(([nhom, items], groupIndex) => {
            html += `<tr class="group-row">
                <td>${groupIndex + 1}</td>
                <td class="text-left">${escapeHtml(`NHÓM ${nhom.toUpperCase()}:`)}</td>
                <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
            </tr>`;

            items.forEach((item, itemIndex) => {
                addNumber(totals, "theoSoSl", item.theoSoSl);
                addNumber(totals, "theoSoNguyenGia", item.theoSoNguyenGia);
                addNumber(totals, "theoSoGiaTri", item.theoSoGiaTri);
                addNumber(totals, "kiemKeSl", item.kiemKeSl);
                addNumber(totals, "kiemKeNguyenGia", item.kiemKeNguyenGia);
                addNumber(totals, "kiemKeGiaTri", item.kiemKeGiaTri);
                addNumber(totals, "chenhLechSl", item.chenhLechSl);
                addNumber(totals, "chenhLechNguyenGia", item.chenhLechNguyenGia);
                addNumber(totals, "chenhLechGiaTri", item.chenhLechGiaTri);

                html += `<tr class="child-row">
                    <td>${groupIndex + 1}.${itemIndex + 1}</td>
                    <td class="text-left">${escapeHtml(item.ten)}</td>
                    <td>${escapeHtml(item.ma)}</td>
                    <td class="text-right">${escapeHtml(formatNum(toNumberOrNull(item.theoSoSl)))}</td>
                    <td class="text-right">${escapeHtml(formatNum(toNumberOrNull(item.theoSoNguyenGia)))}</td>
                    <td class="text-right">${escapeHtml(formatNum(toNumberOrNull(item.theoSoGiaTri)))}</td>
                    <td class="text-right">${escapeHtml(formatNum(toNumberOrNull(item.kiemKeSl)))}</td>
                    <td class="text-right">${escapeHtml(formatNum(toNumberOrNull(item.kiemKeNguyenGia)))}</td>
                    <td class="text-right">${escapeHtml(formatNum(toNumberOrNull(item.kiemKeGiaTri)))}</td>
                    <td class="text-right">${escapeHtml(formatNum(toNumberOrNull(item.chenhLechSl)))}</td>
                    <td class="text-right">${escapeHtml(formatNum(toNumberOrNull(item.chenhLechNguyenGia)))}</td>
                    <td class="text-right">${escapeHtml(formatNum(toNumberOrNull(item.chenhLechGiaTri)))}</td>
                    <td class="text-left">${escapeHtml(item.ghiChu)}</td>
                </tr>`;
            });
        });

        html += `<tr class="grand-total-row">
            <td></td>
            <td class="text-left">TỔNG CỘNG</td>
            <td></td>
            <td class="text-right">${escapeHtml(formatNum(totals.theoSoSl))}</td>
            <td class="text-right">${escapeHtml(formatNum(totals.theoSoNguyenGia))}</td>
            <td class="text-right">${escapeHtml(formatNum(totals.theoSoGiaTri))}</td>
            <td class="text-right">${escapeHtml(formatNum(totals.kiemKeSl))}</td>
            <td class="text-right">${escapeHtml(formatNum(totals.kiemKeNguyenGia))}</td>
            <td class="text-right">${escapeHtml(formatNum(totals.kiemKeGiaTri))}</td>
            <td class="text-right">${escapeHtml(formatNum(totals.chenhLechSl))}</td>
            <td class="text-right">${escapeHtml(formatNum(totals.chenhLechNguyenGia))}</td>
            <td class="text-right">${escapeHtml(formatNum(totals.chenhLechGiaTri))}</td>
            <td></td>
        </tr>`;

        tbody.innerHTML = html;
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
            const errText = await response.text();
            throw new Error(`${tableName}: ${response.status} ${errText}`);
        }
        const data = await response.json();
        if (Array.isArray(data)) return data;
        if (data && Array.isArray(data.Rows)) return data.Rows;
        return [];
    }

    function collectPreviewKeys(rows) {
        const keySet = new Set();
        for (const row of rows || []) {
            if (!row || typeof row !== "object") continue;
            Object.keys(row).forEach((k) => keySet.add(k));
        }
        return [...keySet].sort((a, b) => String(a).localeCompare(String(b), "vi"));
    }

    function previewCellValue(value) {
        if (value == null) return "";
        const text = cellString(value).trim();
        if (text) return text.length > 800 ? `${text.slice(0, 800)}...` : text;
        try {
            const json = JSON.stringify(value);
            if (json) return json.length > 800 ? `${json.slice(0, 800)}...` : json;
        } catch (_) {}
        return String(value);
    }

    function buildRawPreviewTableHtml(rows) {
        if (!rows || rows.length === 0) return '<p class="kktb-prev-note">Không có dòng.</p>';
        const showRows = rows.slice(0, PREVIEW_ROW_CAP);
        const keys = collectPreviewKeys(rows);
        if (!keys.length) return `<p class="kktb-prev-note">${rows.length} dòng - không có khóa cột.</p>`;
        const head = `<tr>${keys.map((key) => `<th>${escapeHtml(key)}</th>`).join("")}</tr>`;
        const body = showRows
            .map((row) => `<tr>${keys.map((key) => `<td>${escapeHtml(previewCellValue(row[key]))}</td>`).join("")}</tr>`)
            .join("");
        const note =
            (rows.length > PREVIEW_ROW_CAP ? `Hiển thị tối đa ${PREVIEW_ROW_CAP}/${rows.length} dòng. ` : `${rows.length} dòng. `) +
            "Kéo ngang nếu bảng rộng.";
        return `<p class="kktb-prev-note">${escapeHtml(note)}</p><div class="kktb-prev-scroll"><table class="kktb-prev-table"><thead>${head}</thead><tbody>${body}</tbody></table></div>`;
    }

    function closeLoadedDataModal() {
        const modal = document.getElementById("kktb-loaded-data-modal");
        if (!modal) return;
        modal.style.display = "none";
        modal.setAttribute("aria-hidden", "true");
        document.body.style.overflow = "";
    }

    function openLoadedDataModal() {
        const modal = document.getElementById("kktb-loaded-data-modal");
        const panel = document.getElementById("kktb-loaded-panel");
        if (!modal || !panel) return;
        if (cacheRows == null) {
            alert("Chưa tải dữ liệu AppSheet. Bấm «Tải dữ liệu AppSheet» trước.");
            return;
        }
        panel.innerHTML = buildRawPreviewTableHtml(cacheRows);
        modal.style.display = "flex";
        modal.setAttribute("aria-hidden", "false");
        document.body.style.overflow = "hidden";
    }

    async function loadFromAppSheet(forceRefresh) {
        const loadBtn = document.getElementById("kktb-btn-load");
        const refreshBtn = document.getElementById("kktb-btn-refresh");
        if (loadBtn) loadBtn.disabled = true;
        if (refreshBtn) refreshBtn.disabled = true;
        try {
            if (forceRefresh) cacheRows = null;
            if (!cacheRows) {
                setStatus(`Đang tải «${TABLE_NXLC_CT}» từ AppSheet...`, false);
                cacheRows = await fetchAppSheetTable(TABLE_NXLC_CT);
            }
            renderReport(cacheRows);
            setStatus(`Đã tải ${cacheRows.length} dòng từ bảng «${TABLE_NXLC_CT}».`, false);
        } catch (error) {
            console.error(error);
            setStatus(`Lỗi tải dữ liệu: ${error.message || error}`, true);
        } finally {
            if (loadBtn) loadBtn.disabled = false;
            if (refreshBtn) refreshBtn.disabled = false;
        }
    }

    function init() {
        document.getElementById("kktb-btn-load")?.addEventListener("click", () => loadFromAppSheet(false));
        document.getElementById("kktb-btn-refresh")?.addEventListener("click", () => loadFromAppSheet(true));
        document.getElementById("kktb-btn-view-loaded")?.addEventListener("click", openLoadedDataModal);
        document.getElementById("kktb-loaded-close")?.addEventListener("click", closeLoadedDataModal);
        document.getElementById("kktb-loaded-data-modal")?.addEventListener("click", (event) => {
            if (event.target && event.target.id === "kktb-loaded-data-modal") closeLoadedDataModal();
        });
        document.addEventListener("keydown", (event) => {
            if (event.key === "Escape") closeLoadedDataModal();
        });
        loadFromAppSheet(false);
    }

    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
    else init();
})();
