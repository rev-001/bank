/**
 * Frontend client logic for Offline PDF-to-Excel Agent
 */

let currentDoc = null;
let currentHeaders = [];
let currentRows = [];
let currentKeyValues = {};

document.addEventListener("DOMContentLoaded", () => {
    initDropzone();
    loadSamples();
    loadStats();
});

// -------------------------------------------------------------
// Dropzone and Ingestion
// -------------------------------------------------------------
function initDropzone() {
    const dropzone = document.getElementById("dropzone");
    const fileInput = document.getElementById("file-input");

    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add("dragover");
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove("dragover");
        });
    });

    dropzone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0 && files[0].type === "application/pdf") {
            uploadPDF(files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (fileInput.files.length > 0) {
            uploadPDF(fileInput.files[0]);
        }
    });
}

async function uploadPDF(file) {
    appendChatMsg("user", `Uploaded: ${file.name}`);
    appendChatMsg("agent", `Ingesting document, matching structural patterns from history, and validating exact calculations...`);

    const formData = new FormData();
    formData.append("file", file);

    try {
        const response = await fetch("/api/extract", {
            method: "POST",
            body: formData
        });
        const data = await response.json();
        if (data.error) {
            appendChatMsg("agent", `Error parsing PDF: ${data.error}`);
            return;
        }
        handleExtractionSuccess(data, file.name);
    } catch (err) {
        appendChatMsg("agent", `Extraction failed: ${err.message}`);
    }
}

async function loadSamples() {
    try {
        const response = await fetch("/api/samples");
        const samples = await response.json();
        const container = document.getElementById("sample-list");
        container.innerHTML = "";

        samples.forEach(s => {
            const card = document.createElement("div");
            card.className = "sample-card";
            card.onclick = () => selectSample(s.path, s.name);
            card.innerHTML = `
                <div>
                    <div class="sample-name">${s.name}</div>
                    <div class="sample-type">${(s.size / 1024).toFixed(1)} KB PDF</div>
                </div>
                <button class="btn-xs">Extract</button>
            `;
            container.appendChild(card);
        });
    } catch (e) {
        console.error("Failed to load samples", e);
    }
}

async function selectSample(samplePath, sampleName) {
    appendChatMsg("user", `Extract sample: ${sampleName}`);
    appendChatMsg("agent", `Analyzing layout and checking historical knowledge base for ${sampleName}...`);

    try {
        const response = await fetch("/api/extract", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ sample_path: samplePath })
        });
        const data = await response.json();
        if (data.error) {
            appendChatMsg("agent", `Error: ${data.error}`);
            return;
        }
        handleExtractionSuccess(data, sampleName);
    } catch (err) {
        appendChatMsg("agent", `Error: ${err.message}`);
    }
}

function handleExtractionSuccess(data, filename) {
    currentDoc = data;
    currentHeaders = data.headers || [];
    currentRows = data.rows || [];
    currentKeyValues = data.key_values || {};

    // Update active doc card
    document.getElementById("active-doc-card").style.display = "block";
    document.getElementById("active-doc-name").innerText = filename;
    document.getElementById("doc-accuracy-badge").innerText = `${(data.accuracy_score * 100).toFixed(1)}% Exact`;
    document.getElementById("doc-template-tag").innerText = data.template_used || "Auto-Discovered";
    document.getElementById("doc-rows-tag").innerText = `${data.row_count} Extracted Rows`;

    // Update download button
    const dlBtn = document.getElementById("btn-download-excel");
    dlBtn.href = data.download_url;
    const outputName = (data.download_url || "").split("/").pop() || "Excel file";
    dlBtn.textContent = `Download ${outputName}`;
    dlBtn.classList.remove("disabled");

    document.getElementById("btn-save-tpl").disabled = false;

    // Render Table
    renderTable(currentHeaders, currentRows, data.is_bank_statement);
    renderKeyValues(currentKeyValues);
    renderAudit(data);

    // Update Tabs Counts
    document.getElementById("tab-row-count").innerText = currentRows.length;
    document.getElementById("tab-kv-count").innerText = Object.keys(currentKeyValues).length;

    // Agent Chat Response
    const exportMessage = data.is_bank_statement
        ? `The clean transaction workbook <code>${escapeHtml(outputName)}</code> uses ₹ currency formatting.`
        : `All transactions are in the combined Excel file <code>all_transactions.xlsx</code>.`;
    appendChatMsg("agent", `Extracted <strong>${data.row_count} transaction rows</strong> with <strong>${(data.accuracy_score * 100).toFixed(1)}% confidence</strong> using <em>${data.template_used}</em>. ${exportMessage}`);
    loadStats();
}

// -------------------------------------------------------------
// Table Rendering & Interactive Cell Editing
// -------------------------------------------------------------
function renderTable(headers, rows, isBankStatement = false) {
    const container = document.getElementById("table-container");
    document.getElementById("table-toolbar").style.display = "flex";

    if (!headers || headers.length === 0) {
        container.innerHTML = `<div class="empty-state"><p class="empty-sub">No tabular data found.</p></div>`;
        return;
    }

    let html = `<table class="excel-table" id="editable-excel-table"><thead><tr>`;
    headers.forEach((h, idx) => {
        html += `<th>${escapeHtml(h)}</th>`;
    });
    html += `</tr></thead><tbody>`;

    rows.forEach((r, r_idx) => {
        html += `<tr data-row="${r_idx}">`;
        r.forEach((cell, c_idx) => {
            const isCurrency = isBankStatement && ["debit", "credit", "balance"].includes(String(headers[c_idx]).trim().toLowerCase());
            const isNum = typeof cell === "number" || (cell !== "" && cell !== null && !isNaN(cell));
            const alignClass = isNum ? "cell-num" : "";
            let displayVal = cell;
            if (isCurrency && cell !== "" && cell !== null) {
                const numericValue = typeof cell === "number" ? cell : Number(String(cell).replace(/[₹,\s]/g, ""));
                displayVal = Number.isFinite(numericValue)
                    ? `₹${numericValue.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
                    : cell;
            } else if (typeof cell === "number") {
                displayVal = cell.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            }
            html += `<td class="${alignClass}" contenteditable="true" data-row="${r_idx}" data-col="${c_idx}">${escapeHtml(displayVal)}</td>`;
        });
        html += `</tr>`;
    });

    html += `</tbody></table>`;
    container.innerHTML = html;
}

function renderKeyValues(kv) {
    const container = document.getElementById("kv-container");
    if (!kv || Object.keys(kv).length === 0) {
        container.innerHTML = `<div class="empty-state"><p class="empty-sub">No document metadata extracted.</p></div>`;
        return;
    }

    let html = "";
    for (const [key, val] of Object.entries(kv)) {
        html += `
            <div class="kv-card">
                <span class="kv-label">${escapeHtml(key)}</span>
                <span class="kv-value">${escapeHtml(val)}</span>
            </div>
        `;
    }
    container.innerHTML = html;
}

function renderAudit(data) {
    const container = document.getElementById("audit-container");
    let html = `
        <div class="audit-card">
            <h4>Mathematical Checksum & Precision Verification</h4>
            <div class="audit-row">
                <span>Extraction Confidence Score</span>
                <span class="audit-status-badge pass">${(data.accuracy_score * 100).toFixed(1)}% Guaranteed</span>
            </div>
            <div class="audit-row">
                <span>Template & Rules Matched</span>
                <span>${escapeHtml(data.template_used)}</span>
            </div>
            <div class="audit-row">
                <span>Applied Historical Rules</span>
                <span>${data.rules_applied_count || 0} active chat/history instructions</span>
            </div>
        </div>
    `;

    if (data.validation_audit && data.validation_audit.length > 0) {
        html += `
            <div class="audit-card">
                <h4>Reconciliation Action Log</h4>
        `;
        data.validation_audit.forEach(item => {
            const isPass = item.type.includes("success") || item.type.includes("autofill") || item.type.includes("correction");
            html += `
                <div class="audit-row">
                    <span>${escapeHtml(item.check || item.type.replace(/_/g, ' '))}</span>
                    <span class="audit-status-badge ${isPass ? 'pass' : 'warn'}">${escapeHtml(item.value || item.reason || 'Verified')}</span>
                </div>
            `;
        });
        html += `</div>`;
    }

    container.innerHTML = html;
}

function addTableRow() {
    if (!currentHeaders || currentHeaders.length === 0) return;
    const newRow = currentHeaders.map(() => "");
    currentRows.push(newRow);
    renderTable(currentHeaders, currentRows, currentDoc && currentDoc.is_bank_statement);
    document.getElementById("tab-row-count").innerText = currentRows.length;
}

async function syncCustomEdits() {
    if (!currentDoc) return;
    const table = document.getElementById("editable-excel-table");
    if (!table) return;

    const rows = [];
    const trs = table.querySelectorAll("tbody tr");
    trs.forEach(tr => {
        const row = [];
        tr.querySelectorAll("td").forEach(td => {
            row.push(td.innerText.trim());
        });
        rows.push(row);
    });

    currentRows = rows;
    appendChatMsg("user", "Exporting updated table with manual edits...");

    try {
        const response = await fetch("/api/export-custom", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                filename: currentDoc.filename,
                headers: currentHeaders,
                rows: currentRows,
                key_values: currentKeyValues,
                is_bank_statement: currentDoc.is_bank_statement
            })
        });
        const res = await response.json();
        if (res.status === "success") {
            const dlBtn = document.getElementById("btn-download-excel");
            dlBtn.href = res.download_url;
            appendChatMsg("agent", "✓ Updated Excel workbook generated with your custom modifications!");
        }
    } catch (e) {
        appendChatMsg("agent", `Error saving edits: ${e.message}`);
    }
}

// -------------------------------------------------------------
// Agent Chat Loop & Learning
// -------------------------------------------------------------
async function handleChatSubmit(e) {
    e.preventDefault();
    const input = document.getElementById("chat-input");
    const prompt = input.value.trim();
    if (!prompt) return;

    input.value = "";
    sendChatPrompt(prompt);
}

function sendQuickPrompt(promptText) {
    sendChatPrompt(promptText);
}

async function sendChatPrompt(prompt) {
    appendChatMsg("user", prompt);

    try {
        const response = await fetch("/api/chat", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                prompt: prompt,
                pdf_path: currentDoc ? currentDoc.pdf_path : null
            })
        });
        const data = await response.json();
        if (data.learning && data.learning.message) {
            appendChatMsg("agent", `✓ ${data.learning.message}`);
        }

        // If active PDF was re-processed with new rules, update live view
        if (data.reprocessed) {
            handleExtractionSuccess(data.reprocessed, currentDoc.filename);
            appendChatMsg("agent", `Re-applied new rules to ${currentDoc.filename} immediately.`);
        }
        loadStats();
    } catch (e) {
        appendChatMsg("agent", `Chat error: ${e.message}`);
    }
}

function appendChatMsg(sender, htmlContent) {
    const feed = document.getElementById("chat-feed");
    const div = document.createElement("div");
    div.className = `chat-msg ${sender}`;
    div.innerHTML = `<div class="msg-bubble">${htmlContent}</div>`;
    feed.appendChild(div);
    feed.scrollTop = feed.scrollHeight;
}

// -------------------------------------------------------------
// Template & History Modals
// -------------------------------------------------------------
function switchTab(tabId) {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));

    event.currentTarget.classList.add("active");
    document.getElementById(tabId).classList.add("active");
}

async function promptSaveTemplate() {
    if (!currentDoc) return;
    const name = prompt("Enter a name for this layout template (e.g. 'Vendor ABC Standard'):");
    if (!name) return;

    try {
        const response = await fetch("/api/save-template", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                template_name: name,
                pdf_path: currentDoc.pdf_path
            })
        });
        const data = await response.json();
        if (data.status === "success") {
            appendChatMsg("agent", `✓ Saved template <strong>${name}</strong> into persistent agent memory.`);
            loadStats();
        }
    } catch (e) {
        alert("Failed to save template: " + e.message);
    }
}

async function loadStats() {
    try {
        const [histRes, tplRes] = await Promise.all([
            fetch("/api/history"),
            fetch("/api/templates")
        ]);
        const history = await histRes.json();
        const templates = await tplRes.json();

        document.getElementById("stat-history-count").innerText = history.length;
        document.getElementById("stat-template-count").innerText = templates.length;
    } catch (e) {
        console.error(e);
    }
}

async function openHistoryModal() {
    const modal = document.getElementById("modal-history");
    const body = document.getElementById("history-modal-body");
    modal.classList.add("active");
    body.innerHTML = "Loading history...";

    try {
        const response = await fetch("/api/history");
        const list = await response.json();
        if (list.length === 0) {
            body.innerHTML = "<p>No previous runs in history.</p>";
            return;
        }

        let html = `<table class="excel-table"><thead><tr><th>ID</th><th>PDF File</th><th>Template</th><th>Rows</th><th>Accuracy</th><th>Date</th></tr></thead><tbody>`;
        list.forEach(item => {
            html += `
                <tr>
                    <td>#${item.id}</td>
                    <td><strong>${escapeHtml(item.pdf_filename)}</strong></td>
                    <td>${escapeHtml(item.template_name || "Auto")}</td>
                    <td class="cell-num">${item.table_row_count}</td>
                    <td class="cell-num">${(item.accuracy_score * 100).toFixed(1)}%</td>
                    <td>${item.created_at.substring(0, 16)}</td>
                </tr>
            `;
        });
        html += `</tbody></table>`;
        body.innerHTML = html;
    } catch (e) {
        body.innerHTML = `<p>Error loading history: ${e.message}</p>`;
    }
}

async function openTemplatesModal() {
    const modal = document.getElementById("modal-templates");
    const body = document.getElementById("templates-modal-body");
    modal.classList.add("active");
    body.innerHTML = "Loading templates...";

    try {
        const response = await fetch("/api/templates");
        const list = await response.json();
        if (list.length === 0) {
            body.innerHTML = "<p>No custom templates saved yet. Templates are automatically learned as you use the agent.</p>";
            return;
        }

        let html = `<table class="excel-table"><thead><tr><th>Template Name</th><th>Headers Recognized</th><th>Usage Count</th></tr></thead><tbody>`;
        list.forEach(t => {
            html += `
                <tr>
                    <td><strong>${escapeHtml(t.name)}</strong></td>
                    <td>${escapeHtml(t.header_patterns.join(", "))}</td>
                    <td class="cell-num">${t.usage_count}</td>
                </tr>
            `;
        });
        html += `</tbody></table>`;
        body.innerHTML = html;
    } catch (e) {
        body.innerHTML = `<p>Error loading templates: ${e.message}</p>`;
    }
}

function closeModal(modalId) {
    document.getElementById(modalId).classList.remove("active");
}

function escapeHtml(text) {
    if (text === null || text === undefined) return "";
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
