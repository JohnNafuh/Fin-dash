// ========================================
// FIN-DASH
// BANK STATEMENT ANALYZER + LIVE MARKET
// ========================================

// ========================================
// ELEMENTS
// ========================================

const analyzer = document.querySelector(".statement-analyzer");
const uploadButton = document.getElementById("chooseStatement");
const statementInput = document.getElementById("statementFile");

const uploadState = document.getElementById("statementUpload");
const analyzingState = document.getElementById("statementAnalyzing");
const completeState = document.getElementById("statementComplete");

// ========================================
// FILE UPLOAD
// ========================================

if (uploadButton && statementInput) {

    uploadButton.addEventListener("click", () => {
        statementInput.click();
    });

    statementInput.addEventListener("change", () => {

        const file = statementInput.files[0];

        if (file) {
            startAnalysis(file);
        }

    });
}

// ========================================
// ANALYZE ANOTHER
// ========================================

document.getElementById("analyzeAnother")?.addEventListener("click", () => {

    statementInput.value = "";

    uploadState.style.display = "block";
    analyzingState.style.display = "none";
    completeState.style.display = "none";

});

// ========================================
// START ANALYSIS
// ========================================

function startAnalysis(file) {

    uploadState.style.display = "none";
    completeState.style.display = "none";
    analyzingState.style.display = "flex";

    const fileName =
        document.getElementById("analyzingFileName");

    const progress =
        document.getElementById("analysisProgressBar");

    if (fileName) {
        fileName.textContent = file.name;
    }

    if (progress) {
        progress.style.width = "0%";
    }

    let value = 0;

    const loading = setInterval(() => {

        value += Math.random() * 8;

        if (value > 90) {
            value = 90;
        }

        if (progress) {
            progress.style.width = value + "%";
        }

    }, 250);

    analyzeStatement(file)

        .then(data => {

            clearInterval(loading);

            if (progress) {
                progress.style.width = "100%";
            }

            setTimeout(() => {

                updateDashboard(data);
                showComplete(data, file.name);

            }, 500);

        })

        .catch(error => {

            clearInterval(loading);

            console.error("Fin-dash analysis error:", error);

            showUploadError();

        });
}

// ========================================
// READ STATEMENT
// ========================================

async function analyzeStatement(file) {

    const extension = file.name.split(".").pop().toLowerCase();

    // CSV / XLSX / XLS -> grid of cells (SheetJS handles quoted commas)
    if (["csv", "xlsx", "xls"].includes(extension)) {

        await loadXLSX();

        const workbook = extension === "csv"
            ? XLSX.read(await file.text(), { type: "string", raw: true })
            : XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });

        let rows = [];

        workbook.SheetNames.forEach(name => {
            rows.push(...XLSX.utils.sheet_to_json(workbook.Sheets[name], {
                header: 1,
                defval: "",
                raw: false
            }));
        });

        return parseRows(rows);
    }

    // PDF -> positioned text items per page
    if (extension === "pdf") {

        await loadPDFJS();

        const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;

        const pages = [];

        for (let pageNo = 1; pageNo <= pdf.numPages; pageNo++) {

            const page = await pdf.getPage(pageNo);
            const content = await page.getTextContent();

            pages.push({
                items: content.items
                    .filter(item => item.str && item.str.trim())
                    .map(item => ({
                        str: item.str.trim(),
                        x: item.transform[4],
                        y: item.transform[5],
                        w: item.width
                    }))
            });
        }

        const hasText = pages.some(page => page.items.length);

        if (!hasText) {
            throw new Error("This PDF has no text layer (it may be a scanned image).");
        }

        return parsePDFStatement(pages);
    }

    throw new Error("Unsupported file type");
}

// ========================================
// COLUMN ROLES (works across banks)
// Kuda, OPay, GTBank, Moniepoint, Access,
// Zenith, PalmPay, UBA, First Bank, etc.
// ========================================

const COLUMN_ROLES = [
    ["date", /^(trans(action)?\.?\s*(date|time)|date\s*\/?\s*time|date|txn\.?\s*date|tran\.?\s*date|posting\s*date|post\s*date|entry\s*date)$/i],
    ["credit", /^(money\s*in|credits?|deposits?|inflow|lodgements?|cr\.?\s*amount|credit\s*amount)$/i],
    ["debit", /^(money\s*out|debits?|withdrawals?|outflow|dr\.?\s*amount|debit\s*amount)$/i],
    ["amount", /^(amount|transaction\s*amount|amt)$/i],
    ["balance", /^(balance|running\s*balance|bal\.?|balance\s*after|available\s*balance|closing\s*balance)/i],
    ["desc", /^(description|narration|narrative|remarks?|details|particulars|to\s*\/\s*from|beneficiary|counterparty|transaction\s*details|memo)$/i],
    ["kind", /^(category|type|transaction\s*type|dr\s*\/\s*cr|cr\s*\/\s*dr)$/i],
    ["ignore", /^(value\s*date|channel|reference|ref\.?|ref\.?\s*no\.?|transaction\s*reference|session\s*id|originating\s*branch|branch|cheque\s*no\.?|time)$/i]
];

function columnRole(text) {

    const clean = String(text || "")
        .replace(/[()₦]/g, " ")
        .replace(/\(?NGN\)?/gi, " ")
        .replace(/\s+/g, " ")
        .trim();

    if (!clean || clean.length > 40) {
        return null;
    }

    for (const [role, pattern] of COLUMN_ROLES) {
        if (pattern.test(clean)) {
            return role;
        }
    }

    return null;
}

// ========================================
// DATES
// ========================================

const MONTHS = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
    jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
};

function parseDate(value) {

    if (value instanceof Date && !isNaN(value)) {
        return value;
    }

    const text = String(value || "").trim();
    let match;

    // 22/08/25, 22-08-2025, 22.08.2025 (day first, Nigerian format)
    if ((match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})\b/))) {
        return buildDate(match[3], Number(match[2]) - 1, match[1], text);
    }

    // 31 Aug 2026, 30-Sep-2023, 01 September, 2025
    if ((match = text.match(/^(\d{1,2})(?:st|nd|rd|th)?[\s-]+([A-Za-z]{3})[A-Za-z]*\.?,?[\s-]+(\d{2,4})\b/))) {
        const month = MONTHS[match[2].toLowerCase()];
        return month === undefined ? null : buildDate(match[3], month, match[1], text);
    }

    // Aug 31, 2026
    if ((match = text.match(/^([A-Za-z]{3})[A-Za-z]*\.?\s+(\d{1,2}),?\s+(\d{4})\b/))) {
        const month = MONTHS[match[1].toLowerCase()];
        return month === undefined ? null : buildDate(match[3], month, match[2], text);
    }

    // 2026-08-31
    if ((match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\b/))) {
        return buildDate(match[1], Number(match[2]) - 1, match[3], text);
    }

    return null;
}

function buildDate(year, month, day, fullText) {

    let y = Number(year);

    if (y < 100) {
        y += 2000;
    }

    const date = new Date(y, month, Number(day));

    if (isNaN(date) || date.getMonth() !== month) {
        return null;
    }

    const time = String(fullText || "").match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);

    if (time) {
        date.setHours(Number(time[1]), Number(time[2]), Number(time[3] || 0));
    }

    return date;
}

function formatDate(date) {

    if (!(date instanceof Date) || isNaN(date)) {
        return "";
    }

    return date.toLocaleDateString("en-NG", {
        day: "2-digit",
        month: "short",
        year: "numeric"
    });
}

// ========================================
// MONEY
// ========================================

// Amounts must have 2 decimals, so account
// numbers and references are never read as money.
const AMOUNT_PATTERN =
    /^[-(]?\s*(?:₦|NGN|N)?\s*-?\s*(?:\d{1,3}(?:,\d{3})+|\d+)?\.\d{2}\s*\)?\s*(?:CR|DR)?$/i;

function isAmount(text) {
    return AMOUNT_PATTERN.test(String(text || "").trim());
}

function parseMoney(value) {

    if (value === undefined || value === null) {
        return 0;
    }

    if (typeof value === "number") {
        return value;
    }

    const raw = String(value).trim();

    if (!raw || /^[-–—]+$/.test(raw)) {
        return 0;
    }

    const negative =
        /^-|^\(|-\s*\d|\bDR$/i.test(raw);

    const number = parseFloat(raw.replace(/[^\d.]/g, ""));

    if (isNaN(number)) {
        return 0;
    }

    return negative ? -Math.abs(number) : number;
}

// ========================================
// PDF: LINES + HEADER DETECTION
// ========================================

function groupLines(items, tolerance = 2.5) {

    const lines = [];

    [...items]
        .sort((a, b) => b.y - a.y)
        .forEach(item => {

            let line = lines.find(existing => Math.abs(existing.y - item.y) <= tolerance);

            if (!line) {
                line = { y: item.y, items: [] };
                lines.push(line);
            }

            line.items.push(item);
        });

    lines.forEach(line => line.items.sort((a, b) => a.x - b.x));

    return lines;
}

function detectHeader(items) {

    // Merge fragments that sit side by side ("Money" + "In", "Debit" + "(₦)")
    const cells = [];

    [...items]
        .sort((a, b) => a.x - b.x)
        .forEach(item => {

            const previous = cells[cells.length - 1];

            if (
                previous &&
                item.x - (previous.x + previous.w) < 4 &&
                Math.abs(item.y - previous.y) < 3
            ) {
                previous.str += " " + item.str;
                previous.w = item.x + item.w - previous.x;
            } else {
                cells.push({ ...item });
            }
        });

    const columns = [];

    cells.forEach(cell => {

        const role = columnRole(cell.str);

        if (role) {
            columns.push({
                role,
                x: cell.x,
                w: cell.w,
                center: cell.x + cell.w / 2
            });
        }
    });

    const has = role => columns.some(column => column.role === role);

    if (has("date") && (has("debit") || has("credit") || has("amount"))) {
        return columns;
    }

    return null;
}

function nearestColumn(item, columns) {

    // Left-aligned tables: item starts exactly under the header
    const aligned = columns.find(column => Math.abs(column.x - item.x) < 3);

    if (aligned) {
        return aligned;
    }

    // Centered / right-aligned tables: closest header center
    const center = item.x + item.w / 2;

    let best = null;
    let bestDistance = Infinity;

    columns.forEach(column => {

        const distance = Math.abs(column.center - center);

        if (distance < bestDistance) {
            bestDistance = distance;
            best = column;
        }
    });

    return best;
}

// ========================================
// PDF: STATEMENT PARSER
// ========================================

function parsePDFStatement(pages) {

    let columns = null;
    const rows = [];

    pages.forEach((page, pageIndex) => {

        // Ignore long paragraph text (footers, disclaimers)
        const items = page.items.filter(item => item.w < 300);
        const lines = groupLines(items);

        const headerYs = [];
        const anchors = [];

        lines.forEach(line => {

            // Header row (also catches headers that wrap onto 2 lines)
            if (detectHeader(line.items)) {

                const band = lines
                    .filter(other => Math.abs(other.y - line.y) <= 8)
                    .flatMap(other => other.items);

                columns = detectHeader(band) || detectHeader(line.items);
                headerYs.push(line.y);

                return;
            }

            if (!columns || headerYs.some(y => Math.abs(line.y - y) <= 8)) {
                return;
            }

            // Every date sitting in the date column starts a new transaction
            const dateColumn = columns.find(column => column.role === "date");

            const dateItem = line.items.find(item =>
                Math.abs(item.x - dateColumn.x) < 25 &&
                parseDate(item.str)
            );

            if (dateItem) {
                anchors.push({
                    y: line.y,
                    page: pageIndex,
                    columns,
                    items: []
                });
            }
        });

        if (!anchors.length) {
            return;
        }

        // How far text can sit from its row (wrapped descriptions, split dates)
        const gaps = anchors
            .slice(1)
            .map((anchor, index) => anchors[index].y - anchor.y)
            .filter(gap => gap > 0)
            .sort((a, b) => a - b);

        const medianGap = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 40;
        const reach = Math.max(12, medianGap * 0.75);

        items.forEach(item => {

            if (headerYs.some(y => Math.abs(item.y - y) <= 8)) {
                return;
            }

            let closest = null;
            let closestDistance = Infinity;

            anchors.forEach(anchor => {

                const distance = Math.abs(anchor.y - item.y);

                if (distance < closestDistance) {
                    closestDistance = distance;
                    closest = anchor;
                }
            });

            if (closest && closestDistance <= reach) {
                closest.items.push(item);
            }
        });

        rows.push(...anchors);
    });

    if (!rows.length) {
        throw new Error("No transaction table found in this PDF.");
    }

    const transactions = [];

    rows.forEach(row => {

        let debit = 0;
        let credit = 0;
        let amount = null;
        let balance = null;

        const dateParts = [];
        const descriptionParts = {};
        const kindParts = [];

        row.items
            .sort((a, b) => b.y - a.y || a.x - b.x)
            .forEach(item => {

                const column = nearestColumn(item, row.columns);

                if (!column) {
                    return;
                }

                const text = item.str;

                if (column.role === "date") {
                    dateParts.push(text);
                    return;
                }

                if (["debit", "credit", "balance", "amount"].includes(column.role)) {

                    if (!isAmount(text)) {
                        return;
                    }

                    const value = parseMoney(text);

                    if (column.role === "debit") debit += Math.abs(value);
                    if (column.role === "credit") credit += Math.abs(value);
                    if (column.role === "balance") balance = value;
                    if (column.role === "amount") amount = (amount || 0) + value;

                    return;
                }

                if (column.role === "desc") {
                    (descriptionParts[column.x] = descriptionParts[column.x] || []).push(text);
                    return;
                }

                if (column.role === "kind") {
                    kindParts.push(text);
                }
            });

        const kind = kindParts.join(" ");

        // Single "Amount" column: sign or type decides direction
        if (amount !== null && !debit && !credit) {

            if (amount < 0 || /\b(dr|debit)\b/i.test(kind)) {
                debit = Math.abs(amount);
            } else {
                credit = Math.abs(amount);
            }
        }

        if (!debit && !credit) {
            return;
        }

        // Keep column order: "To / From" first, then "Description"
        const description = Object.keys(descriptionParts)
            .sort((a, b) => a - b)
            .map(key => descriptionParts[key].join(" "))
            .join(" · ");

        transactions.push(
            buildTransaction(
                parseDate(dateParts.join(" ")),
                description,
                credit ? credit : debit,
                credit ? "income" : "expense",
                kind,
                balance
            )
        );
    });

    return calculateFinancials(transactions, findClosingBalances(pages));
}

// ========================================
// PDF: PRINTED CLOSING BALANCE(S)
// ========================================

function findClosingBalances(pages) {

    const balances = [];

    pages.forEach(page => {

        page.items.forEach(label => {

            if (!/^closing\s*balance:?$/i.test(label.str)) {
                return;
            }

            // Value printed beside the label...
            const beside = page.items.find(item =>
                item !== label &&
                isAmount(item.str) &&
                Math.abs(item.y - label.y) < 3 &&
                item.x > label.x &&
                item.x - (label.x + label.w) < 150
            );

            // ...or just below it
            const below = page.items
                .filter(item =>
                    item !== label &&
                    isAmount(item.str) &&
                    Math.abs(item.x - label.x) < 20 &&
                    label.y - item.y > 0 &&
                    label.y - item.y < 30
                )
                .sort((a, b) => b.y - a.y)[0];

            const value = beside || below;

            if (value) {
                balances.push(parseMoney(value.str));
            }
        });
    });

    return balances;
}

// ========================================
// CSV / XLSX ROW PARSER
// ========================================

function parseRows(rows) {

    rows = rows.filter(row => Array.isArray(row) && row.some(cell => String(cell).trim()));

    if (!rows.length) {
        throw new Error("The file is empty.");
    }

    let headerIndex = -1;
    let roles = [];

    for (let i = 0; i < Math.min(rows.length, 40); i++) {

        const candidate = rows[i].map(cell => columnRole(cell));
        const has = role => candidate.includes(role);

        if (has("date") && (has("debit") || has("credit") || has("amount"))) {
            headerIndex = i;
            roles = candidate;
            break;
        }
    }

    if (headerIndex < 0) {
        throw new Error("Couldn't find the transaction columns.");
    }

    // Use the first matching column of each role
    // (e.g. "Trans. Date" wins over "Value Date")
    const first = role => roles.indexOf(role);

    const dateIndex = first("date");
    const debitIndex = first("debit");
    const creditIndex = first("credit");
    const amountIndex = first("amount");
    const balanceIndex = first("balance");
    const kindIndex = first("kind");

    const descriptionIndexes = roles
        .map((role, index) => (role === "desc" ? index : -1))
        .filter(index => index >= 0);

    const transactions = [];

    for (let i = headerIndex + 1; i < rows.length; i++) {

        const row = rows[i];
        const date = parseDate(row[dateIndex]);

        // Repeated headers, totals, blank lines
        if (!date) {
            continue;
        }

        const debit = debitIndex >= 0 ? Math.abs(parseMoney(row[debitIndex])) : 0;
        const credit = creditIndex >= 0 ? Math.abs(parseMoney(row[creditIndex])) : 0;
        const kind = kindIndex >= 0 ? String(row[kindIndex] || "") : "";

        let type = null;
        let value = 0;

        if (credit > 0) {
            type = "income";
            value = credit;
        } else if (debit > 0) {
            type = "expense";
            value = debit;
        } else if (amountIndex >= 0) {

            const amount = parseMoney(row[amountIndex]);

            if (amount !== 0) {
                // Negative or marked DR/Debit = money out
                const isDebit = amount < 0 || /\b(dr|debit|withdrawal)\b/i.test(kind);
                type = isDebit ? "expense" : "income";
                value = Math.abs(amount);
            }
        }

        if (!type) {
            continue;
        }

        const description = descriptionIndexes
            .map(index => String(row[index] || "").trim())
            .filter(Boolean)
            .join(" · ");

        transactions.push(
            buildTransaction(
                date,
                description,
                value,
                type,
                kind,
                balanceIndex >= 0 ? parseMoney(row[balanceIndex]) : null
            )
        );
    }

    return calculateFinancials(transactions, []);
}

// ========================================
// TRANSACTION BUILDER
// ========================================

function buildTransaction(date, rawDescription, amount, type, kind, balance) {

    const description = cleanDescription(rawDescription);
    const internal = isInternalMove(rawDescription + " " + kind);

    return {
        date,
        dateLabel: formatDate(date),
        description,
        amount: Math.abs(amount),
        type,
        balance,
        internal,
        category: internal
            ? "Savings moves"
            : categorize(rawDescription + " " + kind, type)
    };
}

// ========================================
// INTERNAL MOVES
// Money moving between your own pockets
// (OPay OWealth, Kuda Spend+Save, PalmPay
// CashBox, Piggyvest, etc.) is not real
// income or spending.
// ========================================

function isInternalMove(text) {
    return /owealth\s*withdrawal|auto-?save|spend\s*\+?\s*save|save\s*\+?\s*spend|cashbox|to\s+savings|from\s+savings|savings\s+(top\s*up|withdrawal)|pocket\s+transfer|fixed\s+(deposit|savings)\s+(funding|liquidation)|target\s+savings/i.test(String(text || ""));
}

// ========================================
// CATEGORY
// ========================================

function categorize(text, type) {

    const value = String(text || "").toLowerCase();

    if (type === "income") {
        if (/interest|capitali[sz]ed/.test(value)) return "Interest";
        if (/reversal|refund|cashback/.test(value)) return "Refunds";
        if (/salary|payroll|wages/.test(value)) return "Salary";
        return "Transfers";
    }

    if (/stamp\s*duty|vat|withholding|sms\s*(alert|charge)|maintenance\s*fee|commission|charge|levy|\bfee\b/.test(value)) {
        return "Bills";
    }

    if (/airtime|mobile\s*data|data\s*(plan|bundle|purchase)|\d+(\.\d+)?\s*gb\b|electric|aedc|ikedc|ekedc|phcn|prepaid|water|internet|spectranet|dstv|gotv|startimes|\bbills?\b|utility|recharge|subscription|mtn|airtel|\bglo\b|9mobile/.test(value)) {
        return "Bills";
    }

    if (/netflix|spotify|showmax|apple\.com|youtube\s*premium|movie|cinema|filmhouse|game|betting|bet9ja|sportybet|club|concert|lounge/.test(value)) {
        return "Entertainment";
    }

    if (/food|restaurant|eatery|chicken\s*republic|pizza|grocer|supermarket|shoprite|\bspar\b|foodco|meal|kitchen|cafe|bakery|kfc|domino|chowdeck|glovo|suya/.test(value)) {
        return "Food";
    }

    if (/uber|bolt|taxi|indrive|transport|fuel|petrol|filling\s*station|\bnnpc\b|\bmrs\b|ardova|bus\s*(fare|ticket)|parking|toll/.test(value)) {
        return "Transport";
    }

    if (/jumia|konga|amazon|aliexpress|temu|mall|fashion|clothing|store|stores|\bshop\b|purchase|paystack|flutterwave|korapay|checkout/.test(value)) {
        return "Shopping";
    }

    if (/transfer|trf|tfr|sent\s+to|beneficiary|pos\b|nip/.test(value)) {
        return "Transfers";
    }

    return "Other";
}

// ========================================
// CLEAN DESCRIPTION
// ========================================

function cleanDescription(value) {

    let text = String(value || "");

    text = text
        // Kuda "Name/0123456789/Bank" -> "Name · Bank"
        .replace(/\/\s*\d{6,}\s*\//g, " · ")
        // Masked account numbers 901****233
        .replace(/\b\d{2,4}\*{2,}\d{2,4}\b/g, "")
        // Reference junk (AT139_TRF|..., APITRANSFER-...)
        .replace(/\b[A-Z0-9]+_TRF\|[^\s]*/gi, "")
        .replace(/APITRANSFER-[^\s|]*/gi, "")
        // UUIDs, long numbers, long IDs
        .replace(/\b[0-9a-f]{8}-[0-9a-f-]{8,}\b/gi, "")
        .replace(/\b\d{9,}\b/g, "")
        .replace(/\b[A-Z0-9]{16,}\b/g, "")
        // Dates and amounts that slipped in
        .replace(/\b\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}\b/g, "")
        .replace(/(?:₦|NGN)\s*[\d,]+(?:\.\d{1,2})?/gi, "")
        // Kuda's generic "kuda" description cell
        .replace(/·\s*kuda\s*$/i, "")
        // Separators
        .replace(/\s*\|\s*(\|\s*)*/g, " · ")
        .replace(/(\s*·\s*)+/g, " · ")
        .replace(/^\s*·\s*|\s*·\s*$/g, "")
        .replace(/\s+/g, " ")
        .trim();

    if (!text) {
        return "Bank transaction";
    }

    if (text.length > 55) {
        text = text.slice(0, 55).trim() + "…";
    }

    return text;
}

// ========================================
// CALCULATE FINANCIALS
// ========================================

function emptyCategories() {
    return {
        Food: 0,
        Transport: 0,
        Bills: 0,
        Entertainment: 0,
        Shopping: 0,
        Transfers: 0,
        Other: 0
    };
}

function calculateFinancials(transactions, closingBalances) {

    // Oldest -> newest (statements aren't always in order)
    transactions.sort((a, b) => (a.date || 0) - (b.date || 0));

    let income = 0;
    let expenses = 0;
    let internalVolume = 0;

    const categories = emptyCategories();

    transactions.forEach(transaction => {

        if (transaction.internal) {
            internalVolume += transaction.amount;
            return;
        }

        if (transaction.type === "income") {
            income += transaction.amount;
            return;
        }

        expenses += transaction.amount;

        if (categories[transaction.category] !== undefined) {
            categories[transaction.category] += transaction.amount;
        } else {
            categories.Other += transaction.amount;
        }
    });

    // Balance: printed closing balance(s) first, then last balance column value
    let balance = null;

    if (closingBalances && closingBalances.length) {
        balance = closingBalances.reduce((sum, value) => sum + value, 0);
    } else {
        const withBalance = transactions.filter(t => t.balance !== null && t.balance !== undefined);
        if (withBalance.length) {
            balance = withBalance[withBalance.length - 1].balance;
        }
    }

    const dates = transactions.map(t => t.date).filter(Boolean);

    return {
        transactions,
        income,
        expenses,
        savings: income - expenses,
        balance: balance === null ? income - expenses : balance,
        balanceFromStatement: balance !== null,
        internalVolume,
        categories,
        periodStart: dates.length ? dates[0] : null,
        periodEnd: dates.length ? dates[dates.length - 1] : null
    };
}

function emptyData() {
    return {
        transactions: [],
        income: 0,
        expenses: 0,
        savings: 0,
        balance: 0,
        balanceFromStatement: false,
        internalVolume: 0,
        categories: emptyCategories(),
        periodStart: null,
        periodEnd: null
    };
}

// ========================================
// DASHBOARD UPDATE
// ========================================

function updateDashboard(data) {

    const balance =
        document.querySelector(
            ".balance h2"
        );

    const cards =
        document.querySelectorAll(
            ".summary .card h3"
        );

    // BALANCE
    if (balance) {

        balance.textContent =
            formatNaira(
                data.balance
            );

        balance.title =
            formatNairaExact(
                data.balance
            );
    }

    // SUMMARY CARDS
    if (
        cards.length >= 3
    ) {

        cards[0].textContent =
            formatNaira(
                data.income
            );

        cards[0].title =
            formatNairaExact(
                data.income
            );

        cards[1].textContent =
            formatNaira(
                data.expenses
            );

        cards[1].title =
            formatNairaExact(
                data.expenses
            );

        cards[2].textContent =
            formatNaira(
                data.savings
            );

        cards[2].title =
            formatNairaExact(
                data.savings
            );

    }

    // PERIOD LABELS (replace hardcoded "August" / "This Month")
    const budgetMonth =
        document.querySelector(".budget .section-title span");

    const spendingPeriod =
        document.querySelector(".spending .section-title span");

    if (budgetMonth && data.periodEnd) {
        budgetMonth.textContent =
            data.periodEnd.toLocaleDateString("en-NG", {
                month: "long",
                year: "numeric"
            });
    }

    if (spendingPeriod && data.periodStart && data.periodEnd) {
        spendingPeriod.textContent =
            formatDate(data.periodStart) + " – " + formatDate(data.periodEnd);
    }

    updateTransactions(data);

    updateSpending(data);

    updateBudget(data);
}

// ========================================
// RECENT TRANSACTIONS
// ========================================

function updateTransactions(data) {

    const list =
        document.querySelector(
            ".transaction-list"
        );

    if (!list) {
        return;
    }

    if (
        !data.transactions.length
    ) {

        list.innerHTML = `

            <div class="transaction empty-transaction">

                <p>
                    No transactions found.
                </p>

            </div>

        `;

        return;
    }

    list.innerHTML =
        data.transactions

            .filter(transaction => !transaction.internal)

            .slice(-8)

            .reverse()

            .map(
                transaction => {

                    const exact =
                        formatNairaExact(
                            transaction.amount
                        );

                    return `

                        <div class="transaction">

                            <div>

                                <strong>
                                    ${escapeHTML(
                                        cleanDescription(
                                            transaction.description
                                        )
                                    )}
                                </strong>

                                <small>
                                    ${escapeHTML(
                                        transaction.category
                                    )}${
                                        transaction.dateLabel
                                            ? " · " + escapeHTML(transaction.dateLabel)
                                            : ""
                                    }
                                </small>

                            </div>

                            <strong
                                class="${
                                    transaction.type === "income"
                                        ? "income"
                                        : "expense"
                                }"
                                title="${exact}"
                            >

                                ${
                                    transaction.type === "income"
                                        ? "+"
                                        : "-"
                                }

                                ${formatNaira(
                                    transaction.amount
                                )}

                            </strong>

                        </div>

                    `;

                }
            )
            .join("");
}

// ========================================
// SPENDING OVERVIEW
// ========================================

function updateSpending(data) {

    const items =
        document.querySelectorAll(
            ".spending-item"
        );

    if (!items.length) {
        return;
    }

    const categories = [

        "Food",

        "Transport",

        "Bills",

        "Entertainment"

    ];

    const values =
        categories.map(
            category =>
                Number(
                    data.categories[
                        category
                    ] || 0
                )
        );

    const max =
        Math.max(
            ...values,
            1
        );

    items.forEach(
        (item, index) => {

            const category =
                categories[index];

            if (!category) {
                return;
            }

            const value =
                Number(
                    data.categories[
                        category
                    ] || 0
                );

            const amount =
                item.querySelector(
                    "span"
                );

            const bar =
                item.querySelector(
                    ".progress-bar"
                );

            if (amount) {

                amount.textContent =
                    formatNaira(
                        value
                    );

                amount.title =
                    formatNairaExact(
                        value
                    );

            }

            if (bar) {

                const percentage =
                    value > 0
                        ? (
                            value / max
                        ) * 100
                        : 0;

                bar.style.width =
                    Math.min(
                        percentage,
                        100
                    ) + "%";

            }

        }
    );
}

// ========================================
// MONTHLY BUDGET
// ========================================

function updateBudget(data) {

    const amount =
        document.querySelector(
            ".budget-amount h3"
        );

    const description =
        document.querySelector(
            ".budget-amount p"
        );

    const bar =
        document.querySelector(
            ".budget-progress-bar"
        );

    const message =
        document.querySelector(
            ".budget-message"
        );

    const budget =
        data.income;

    const spent =
        data.expenses;

    const percentage =
        budget > 0
            ? Math.min(
                (
                    spent /
                    budget
                ) * 100,
                100
            )
            : 0;

    if (amount) {

        amount.textContent =
            formatNaira(
                spent
            );

        amount.title =
            formatNairaExact(
                spent
            );
    }

    if (description) {

        description.textContent =
            "of " +
            formatNaira(
                budget
            ) +
            " available";

        description.title =
            formatNairaExact(
                budget
            );

    }

    if (bar) {

        bar.style.width =
            percentage + "%";

    }

    if (message) {

        if (!budget) {

            message.textContent =
                "No income was detected in this statement.";

        } else if (
            spent > budget
        ) {

            message.textContent =
                "Your spending is higher than your recorded income.";

        } else {

            message.textContent =
                Math.round(
                    percentage
                ) +
                "% of your available income has been spent.";

        }

    }
}

// ========================================
// COMPLETE STATE
// ========================================

function showComplete(
    data,
    fileName
) {

    analyzingState.style.display =
        "none";

    completeState.style.display =
        "flex";

    const name =
        document.getElementById(
            "statementFileName"
        );

    const count =
        document.getElementById(
            "transactionCount"
        );

    const income =
        document.getElementById(
            "statementIncome"
        );

    const expenses =
        document.getElementById(
            "statementExpenses"
        );

    if (name) {

        name.textContent =
            fileName;

    }

    if (count) {

        count.textContent =
            data.transactions.length;

    }

    if (income) {

        income.textContent =
            formatNaira(
                data.income
            );

        income.title =
            formatNairaExact(
                data.income
            );

    }

    if (expenses) {

        expenses.textContent =
            formatNaira(
                data.expenses
            );

        expenses.title =
            formatNairaExact(
                data.expenses
            );

    }
}

// ========================================
// ERROR
// ========================================

function showUploadError() {

    analyzingState.style.display =
        "none";

    completeState.style.display =
        "none";

    uploadState.style.display =
        "block";

    const title =
        uploadState.querySelector(
            "h2"
        );

    const message =
        uploadState.querySelector(
            "p"
        );

    if (title) {

        title.textContent =
            "Couldn't analyze statement";

    }

    if (message) {

        message.textContent =
            "Make sure it's a bank statement (CSV, XLSX, XLS or a text-based PDF), not a scanned photo.";

    }

    setTimeout(() => {

        if (title) {

            title.textContent =
                "Upload Bank Statement";

        }

        if (message) {

            message.textContent =
                "Turn your transactions into a clear picture of your finances.";

        }

    }, 5000);
}

// ========================================
// COMPACT NAIRA FORMAT
// ========================================

function formatNaira(value) {

    const number =
        Number(value || 0);

    const absolute =
        Math.abs(number);

    let formatted;

    // MILLIONS
    if (
        absolute >= 1000000
    ) {

        formatted =
            trimCompactDecimal(
                number / 1000000
            ) +
            "M";

    }

    // THOUSANDS
    else if (
        absolute >= 1000
    ) {

        formatted =
            trimCompactDecimal(
                number / 1000
            ) +
            "k";

    }

    // NORMAL
    else {

        formatted =
            number.toLocaleString(
                "en-NG",
                {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2
                }
            );

    }

    return "₦" + formatted;
}

// ========================================
// COMPACT DECIMAL
// ========================================

function trimCompactDecimal(value) {

    return Number(
        value.toFixed(2)
    ).toString();
}

// ========================================
// EXACT NAIRA FORMAT
// ========================================

function formatNairaExact(value) {

    return "₦" +
        Number(value || 0).toLocaleString(
            "en-NG",
            {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            }
        );
}

// ========================================
// SECURITY
// ========================================

function escapeHTML(value) {

    return String(value)

        .replace(
            /&/g,
            "&amp;"
        )

        .replace(
            /</g,
            "&lt;"
        )

        .replace(
            />/g,
            "&gt;"
        )

        .replace(
            /"/g,
            "&quot;"
        )

        .replace(
            /'/g,
            "&#039;"
        );
}

// ========================================
// LOAD XLSX
// ========================================

function loadXLSX() {

    return new Promise(
        (resolve, reject) => {

            if (window.XLSX) {

                resolve();

                return;
            }

            const script =
                document.createElement(
                    "script"
                );

            script.src =
                "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";

            script.onload =
                resolve;

            script.onerror =
                reject;

            document.head.appendChild(
                script
            );

        }
    );
}

// ========================================
// LOAD PDF.JS
// ========================================

function loadPDFJS() {

    return new Promise(
        (resolve, reject) => {

            if (window.pdfjsLib) {

                resolve();

                return;
            }

            const script =
                document.createElement(
                    "script"
                );

            script.src =
                "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";

            script.onload = () => {

                if (window.pdfjsLib) {

                    pdfjsLib
                        .GlobalWorkerOptions
                        .workerSrc =
                        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

                    resolve();

                } else {

                    reject(
                        new Error(
                            "PDF.js failed"
                        )
                    );

                }

            };

            script.onerror =
                reject;

            document.head.appendChild(
                script
            );

        }
    );
}

// ========================================
// LIVE MARKET
// ========================================

const marketBar =
    document.querySelector(
        ".market-wrapper"
    );

const marketToggle =
    document.getElementById(
        "marketToggle"
    );

const marketLoading =
    document.getElementById(
        "marketLoading"
    );

const marketData =
    document.getElementById(
        "marketData"
    );

const canvas =
    document.getElementById(
        "marketChart"
    );

let marketOpen = false;

let priceTimer;

let resizeTimer;

// ========================================
// MARKET TOGGLE
// ========================================

if (marketToggle) {

    marketToggle.addEventListener(
        "click",
        () => {

            marketOpen =
                !marketOpen;

            if (marketOpen) {

                marketBar.style.height =
                    "280px";

                marketToggle.textContent =
                    "⌄";

                marketLoading.style.display =
                    "flex";

                marketData.style.display =
                    "none";

                setTimeout(() => {

                    if (!marketOpen) {
                        return;
                    }

                    marketLoading.style.display =
                        "none";

                    marketData.style.display =
                        "block";

                    drawFinancialChart();

                    startMarketNumbers();

                }, 900);

            } else {

                marketBar.style.height =
                    "55px";

                marketToggle.textContent =
                    "⌃";

                marketLoading.style.display =
                    "none";

                marketData.style.display =
                    "none";

                stopMarket();

            }

        }
    );
}

// ========================================
// MARKET CHART DATA
// ========================================

let marketPoints = [

    48,46,47,44,45,42,43,40,42,39,

    41,37,38,35,36,33,35,32,34,31,

    33,29,31,28,30,27,29,25,27,24,

    26,22,25,21,23,19,22,18,20,17

];

// ========================================
// DRAW MARKET CHART
// ========================================

function drawFinancialChart() {

    if (!canvas) {
        return;
    }

    const rect =
        canvas.getBoundingClientRect();

    const width =
        Math.max(
            rect.width,
            1
        );

    const height =
        Math.max(
            rect.height,
            1
        );

    const dpr =
        window.devicePixelRatio ||
        1;

    canvas.width =
        width * dpr;

    canvas.height =
        height * dpr;

    const ctx =
        canvas.getContext(
            "2d"
        );

    ctx.setTransform(
        dpr,
        0,
        0,
        dpr,
        0,
        0
    );

    ctx.clearRect(
        0,
        0,
        width,
        height
    );

    // GRID
    ctx.strokeStyle =
        "rgba(255,255,255,.055)";

    ctx.lineWidth = 1;

    for (
        let i = 1;
        i <= 4;
        i++
    ) {

        const y =
            (
                height / 5
            ) * i;

        ctx.beginPath();

        ctx.moveTo(
            0,
            y
        );

        ctx.lineTo(
            width,
            y
        );

        ctx.stroke();

    }

    const min =
        Math.min(
            ...marketPoints
        ) - 4;

    const max =
        Math.max(
            ...marketPoints
        ) + 4;

    const range =
        max - min;

    const points =
        marketPoints.map(
            (
                value,
                index
            ) => ({

                x:
                    (
                        index /
                        (
                            marketPoints.length - 1
                        )
                    ) *
                    width,

                y:
                    height -
                    (
                        (
                            value - min
                        ) /
                        range
                    ) *
                    (
                        height - 14
                    ) -
                    7

            })
        );

    function spline() {

        ctx.beginPath();

        ctx.moveTo(
            points[0].x,
            points[0].y
        );

        for (
            let i = 0;
            i < points.length - 1;
            i++
        ) {

            const current =
                points[i];

            const next =
                points[i + 1];

            const midpoint =
                (
                    current.x +
                    next.x
                ) / 2;

            ctx.bezierCurveTo(

                midpoint,
                current.y,

                midpoint,
                next.y,

                next.x,
                next.y

            );

        }

    }

    // AREA
    spline();

    ctx.lineTo(
        width,
        height
    );

    ctx.lineTo(
        0,
        height
    );

    ctx.closePath();

    const gradient =
        ctx.createLinearGradient(
            0,
            0,
            0,
            height
        );

    gradient.addColorStop(
        0,
        "rgba(108,60,255,.25)"
    );

    gradient.addColorStop(
        1,
        "rgba(108,60,255,0)"
    );

    ctx.fillStyle =
        gradient;

    ctx.fill();

    // GLOW
    spline();

    ctx.strokeStyle =
        "rgba(145,112,255,.35)";

    ctx.lineWidth = 7;

    ctx.lineCap =
        "round";

    ctx.lineJoin =
        "round";

    ctx.shadowBlur = 16;

    ctx.shadowColor =
        "rgba(108,60,255,.8)";

    ctx.stroke();

    // MAIN LINE
    spline();

    ctx.shadowBlur = 0;

    ctx.strokeStyle =
        "#bda8ff";

    ctx.lineWidth =
        2.4;

    ctx.stroke();

    // LAST POINT
    const last =
        points[
            points.length - 1
        ];

    ctx.beginPath();

    ctx.arc(
        last.x,
        last.y,
        3.5,
        0,
        Math.PI * 2
    );

    ctx.fillStyle =
        "#fff";

    ctx.shadowBlur = 12;

    ctx.shadowColor =
        "#bda8ff";

    ctx.fill();
}

// ========================================
// MARKET NUMBERS
// ========================================

function startMarketNumbers() {

    clearInterval(
        priceTimer
    );

    const price =
        document.getElementById(
            "usdPrice"
        );

    const change =
        document.getElementById(
            "usdChange"
        );

    if (
        !price ||
        !change
    ) {

        return;
    }

    let currentPrice =
        1530;

    priceTimer =
        setInterval(
            () => {

                currentPrice +=
                    (
                        Math.random() -
                        .46
                    ) * 2;

                price.textContent =
                    "₦" +
                    currentPrice.toLocaleString(
                        "en-NG",
                        {
                            minimumFractionDigits: 2,
                            maximumFractionDigits: 2
                        }
                    );

                change.textContent =
                    "+" +
                    (
                        Math.random() * .7 +
                        .1
                    ).toFixed(2) +
                    "%";

                const last =
                    marketPoints[
                        marketPoints.length - 1
                    ];

                marketPoints.push(
                    last +
                    (
                        Math.random() -
                        .47
                    ) * 5
                );

                if (
                    marketPoints.length >
                    42
                ) {

                    marketPoints.shift();

                }

                drawFinancialChart();

            },
            3000
        );
}

// ========================================
// STOP MARKET
// ========================================

function stopMarket() {

    clearInterval(
        priceTimer
    );

    priceTimer = null;
}

// ========================================
// RESIZE
// ========================================

window.addEventListener(
    "resize",
    () => {

        clearTimeout(
            resizeTimer
        );

        resizeTimer =
            setTimeout(
                () => {

                    if (
                        marketOpen
                    ) {

                        drawFinancialChart();

                    }

                },
                150
            );

    }
);
