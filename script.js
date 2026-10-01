// ========================================
// FIN-DASH
// Bank statement analyzer. Runs entirely in the browser.
// ========================================

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

    const closings = findClosingBalances(pages);
    const allText = pages.flatMap(page => page.items.map(item => item.str)).join(" ");

    return finishStatement(transactions, closings, allText);
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

    const allText = rows.slice(0, headerIndex + 1).flat().join(" ");

    return finishStatement(transactions, [], allText);
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
// STATEMENT RESULT
// ========================================

function finishStatement(transactions, closingBalances, text) {

    transactions.sort((a, b) => (a.date || 0) - (b.date || 0));

    let closing = null;

    if (closingBalances.length) {
        closing = closingBalances.reduce((sum, value) => sum + value, 0);
    } else {
        const withBalance = transactions.filter(t => t.balance !== null && t.balance !== undefined);
        if (withBalance.length) {
            closing = withBalance[withBalance.length - 1].balance;
        }
    }

    return {
        bank: detectBank(text),
        closing,
        transactions
    };
}

function detectBank(text) {

    const value = String(text || "").toLowerCase();

    const banks = [
        ["Kuda", /kuda/],
        ["OPay", /opay digital|opayweb|\bopay\b.*statement|owealth/],
        ["GTBank", /guaranty trust|gtco|gtbank|\bgt\s?crea8|fullstmt/],
        ["Moniepoint", /moniepoint/],
        ["PalmPay", /palmpay/],
        ["Access Bank", /access bank/],
        ["Zenith Bank", /zenith/],
        ["UBA", /united bank for africa|\buba\b/],
        ["First Bank", /first bank|firstbank/],
        ["Wema / ALAT", /wema|alat/],
        ["Fidelity", /fidelity/],
        ["Stanbic IBTC", /stanbic/],
        ["Sterling", /sterling bank/],
        ["FCMB", /\bfcmb\b|first city monument/],
        ["Ecobank", /ecobank/],
        ["Union Bank", /union bank/],
        ["Polaris", /polaris/],
        ["Carbon", /\bcarbon\b/],
        ["FairMoney", /fairmoney/]
    ];

    // Look at the statement header area first (first ~600 chars),
    // because transaction rows mention other banks.
    const head = value.slice(0, 600);

    for (const [name, pattern] of banks) {
        if (pattern.test(head)) return name;
    }

    for (const [name, pattern] of banks) {
        if (pattern.test(value)) return name;
    }

    return "Bank";
}


// ========================================
// APP STATE
// ========================================

const STORE_KEY = "findash:statements:v2";
const BUDGET_KEY = "findash:budget";
const GOAL_KEY = "findash:goal";
const TX_PREVIEW = 5;

const CATEGORY_COLORS = {
    Transfers: "var(--c-transfers)",
    Bills: "var(--c-bills)",
    Food: "var(--c-food)",
    Shopping: "var(--c-shopping)",
    Transport: "var(--c-transport)",
    Entertainment: "var(--c-entertainment)",
    Other: "var(--c-other)",
    Interest: "var(--in)",
    Refunds: "var(--in)",
    Salary: "var(--in)",
    "Savings moves": "var(--ink-3)"
};

const state = {
    statements: [],
    filter: "all",
    search: "",
    shown: TX_PREVIEW,
    summary: null
};

const $ = id => document.getElementById(id);

// ========================================
// STORAGE (safe: works even if blocked)
// ========================================

function readStore(key, fallback) {
    try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : fallback;
    } catch (error) {
        return fallback;
    }
}

function writeStore(key, value) {
    try {
        if (value === null) {
            localStorage.removeItem(key);
        } else {
            localStorage.setItem(key, JSON.stringify(value));
        }
    } catch (error) {
        // Storage full or disabled: the dashboard still works for this visit
    }
}

function saveStatements() {
    writeStore(STORE_KEY, state.statements.map(statement => ({
        ...statement,
        transactions: statement.transactions.map(t => ({
            ...t,
            date: t.date ? t.date.toISOString() : null
        }))
    })));
}

function loadStatements() {
    const saved = readStore(STORE_KEY, []);

    state.statements = Array.isArray(saved)
        ? saved.map(statement => ({
            ...statement,
            transactions: (statement.transactions || []).map(t => {
                const date = t.date ? new Date(t.date) : null;
                return { ...t, date, dateLabel: formatDate(date) };
            })
        }))
        : [];
}

// ========================================
// FILE INPUT + DRAG AND DROP
// ========================================

const fileInput = $("fileInput");

function openPicker() {
    fileInput.value = "";
    fileInput.click();
}

$("chooseBtn").addEventListener("click", event => {
    event.stopPropagation();
    openPicker();
});

$("addBtn").addEventListener("click", openPicker);

$("dropZone").addEventListener("click", openPicker);

$("dropZone").addEventListener("keydown", event => {
    if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openPicker();
    }
});

fileInput.addEventListener("change", () => {
    if (fileInput.files.length) {
        handleFiles([...fileInput.files]);
    }
});

["dragenter", "dragover"].forEach(type => {
    document.addEventListener(type, event => {
        event.preventDefault();
        $("dropZone").classList.add("is-over");
    });
});

["dragleave", "drop"].forEach(type => {
    document.addEventListener(type, event => {
        event.preventDefault();
        if (type === "drop" || event.target === document.documentElement) {
            $("dropZone").classList.remove("is-over");
        }
    });
});

document.addEventListener("drop", event => {
    const files = [...(event.dataTransfer?.files || [])];
    if (files.length) {
        handleFiles(files);
    }
});

// ========================================
// PROCESS FILES
// ========================================

async function handleFiles(files) {

    showView("loading");

    const failures = [];
    let added = 0;

    for (let i = 0; i < files.length; i++) {

        const file = files[i];

        $("loadingText").textContent =
            files.length > 1
                ? `Reading ${file.name} (${i + 1} of ${files.length})…`
                : `Reading ${file.name}…`;

        $("loadingBar").style.width = Math.round((i / files.length) * 100 + 10) + "%";

        try {
            const result = await analyzeStatement(file);

            if (!result.transactions.length) {
                throw new Error("No transactions found.");
            }

            const duplicate = state.statements.some(s =>
                s.name === file.name &&
                s.transactions.length === result.transactions.length
            );

            if (duplicate) {
                failures.push(`${file.name} is already added.`);
                continue;
            }

            state.statements.push({
                id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
                name: file.name,
                bank: result.bank !== "Bank"
                    ? result.bank
                    : (detectBank(file.name) !== "Bank" ? detectBank(file.name) : "Bank account"),
                closing: result.closing,
                transactions: result.transactions
            });

            added++;

        } catch (error) {
            console.error("Fin-dash:", file.name, error);
            failures.push(`${file.name}: ${readableError(error)}`);
        }
    }

    $("loadingBar").style.width = "100%";

    if (added) {
        saveStatements();
    }

    render();

    if (failures.length) {
        toast(failures.join(" "));
    }
}

function readableError(error) {

    const message = String(error && error.message || "");

    if (/no text layer/i.test(message)) {
        return "this PDF is a scanned image. Download the statement again from your bank app as a PDF or CSV.";
    }

    if (/password/i.test(message)) {
        return "this PDF is password-protected. Open it, save a copy without the password, then add that copy.";
    }

    if (/transaction (table|columns)|no transactions/i.test(message)) {
        return "couldn't find a transaction table. Check that it's a bank statement, not a receipt.";
    }

    if (/unsupported/i.test(message)) {
        return "use a PDF, CSV, XLSX or XLS file.";
    }

    if (/load|network|script/i.test(message) || error instanceof Event) {
        return "the file reader couldn't load. Check your internet connection and try again.";
    }

    return "couldn't read this file.";
}

let toastTimer;

function toast(message) {
    const el = $("toast");
    el.textContent = message;
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { el.hidden = true; }, 7000);
}

// ========================================
// VIEWS
// ========================================

function showView(name) {
    $("emptyView").hidden = name !== "empty";
    $("loadingView").hidden = name !== "loading";
    $("dashView").hidden = name !== "dash";
    $("topbarActions").hidden = name !== "dash";
}

// ========================================
// SUMMARY
// ========================================

function summarize() {

    const all = state.statements.flatMap(statement =>
        statement.transactions.map(t => ({ ...t, account: statement.bank }))
    );

    all.sort((a, b) => (a.date || 0) - (b.date || 0));

    const real = all.filter(t => !t.internal);
    const incomeTx = real.filter(t => t.type === "income");
    const expenseTx = real.filter(t => t.type === "expense");

    const sum = list => list.reduce((total, t) => total + t.amount, 0);

    const categories = {};

    expenseTx.forEach(t => {
        categories[t.category] = (categories[t.category] || 0) + t.amount;
    });

    const closings = state.statements
        .map(s => s.closing)
        .filter(value => typeof value === "number" && isFinite(value));

    const dates = all.map(t => t.date).filter(Boolean);

    // Days actually covered by statements (gaps between statements don't count)
    const coveredKeys = new Set();
    const covered = [];

    state.statements.forEach(statement => {
        const list = statement.transactions.map(t => t.date).filter(Boolean).sort((a, b) => a - b);
        if (!list.length) return;
        const d = new Date(list[0].getFullYear(), list[0].getMonth(), list[0].getDate());
        const last = list[list.length - 1];
        for (; d <= last; d.setDate(d.getDate() + 1)) {
            const key = dayKey(d);
            if (!coveredKeys.has(key)) {
                coveredKeys.add(key);
                covered.push(new Date(d));
            }
        }
    });

    covered.sort((a, b) => a - b);

    return {
        covered,
        all,
        real,
        incomeTx,
        expenseTx,
        income: sum(incomeTx),
        expenses: sum(expenseTx),
        internal: sum(all.filter(t => t.internal && t.type === "expense")),
        balance: closings.length ? closings.reduce((a, b) => a + b, 0) : null,
        categories,
        start: dates.length ? dates[0] : null,
        end: dates.length ? dates[dates.length - 1] : null
    };
}

// ========================================
// RENDER
// ========================================

function render() {

    if (!state.statements.length) {
        showView("empty");
        return;
    }

    const summary = summarize();
    state.summary = summary;

    showView("dash");

    renderNote(summary);
    renderFlow(summary);
    renderChart(summary);
    renderBudget(summary);
    renderInsights(summary);
    renderPeople(summary);
    renderTransactions();
    renderGoal(summary);
    renderFiles();
}

// ---------- Balance note ----------

function renderNote(s) {

    const net = s.income - s.expenses;

    $("balanceFigure").textContent =
        s.balance === null ? formatNairaExact(net) : formatNairaExact(s.balance);

    const count = state.statements.length;

    $("periodText").textContent =
        (s.start && s.end ? `${formatDate(s.start)} to ${formatDate(s.end)}, ` : "") +
        `from ${count} statement${count === 1 ? "" : "s"}` +
        (s.balance === null ? ". No closing balance printed, so this shows money in minus money out." : "");

    $("inFigure").textContent = formatNaira(s.income);
    $("inFigure").title = formatNairaExact(s.income);

    $("outFigure").textContent = formatNaira(s.expenses);
    $("outFigure").title = formatNairaExact(s.expenses);

    $("netFigure").textContent = formatNaira(Math.abs(net));
    $("netFigure").title = formatNairaExact(net);

    $("netFigure").parentElement.querySelector("dt").textContent =
        net < 0 ? "You overspent by" : "You kept";
}

// ---------- Where it went ----------

function renderFlow(s) {

    const entries = Object.entries(s.categories)
        .filter(([, value]) => value > 0)
        .sort((a, b) => b[1] - a[1]);

    $("flowMeta").textContent = formatNairaExact(s.expenses);

    if (!entries.length) {
        $("flowStrip").innerHTML = "";
        $("flowLegend").innerHTML = `<li class="muted">No spending in these statements.</li>`;
        return;
    }

    $("flowStrip").innerHTML = entries
        .map(([name, value]) =>
            `<span style="flex:${value};background:${CATEGORY_COLORS[name] || "var(--c-other)"}" title="${escapeHTML(name)}: ${formatNairaExact(value)}"></span>`
        )
        .join("");

    $("flowLegend").innerHTML = entries
        .map(([name, value]) => `
            <li>
                <span class="swatch" style="background:${CATEGORY_COLORS[name] || "var(--c-other)"}"></span>
                <span>${escapeHTML(name)}</span>
                <span class="pct">${Math.round((value / s.expenses) * 100)}%</span>
                <span class="amt" title="${formatNairaExact(value)}">${formatNaira(value)}</span>
            </li>
        `)
        .join("");
}

// ---------- Day by day chart ----------

function dayKey(date) {
    return date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate();
}

function renderChart(s) {

    const chart = $("chart");

    if (!s.start || !s.end) {
        chart.innerHTML = `<p class="muted">No dates found.</p>`;
        return;
    }

    // Latest 90 days that your statements cover
    const visible = s.covered.slice(-90).map(date => ({ date, in: 0, out: 0 }));
    const index = new Map(visible.map((day, i) => [dayKey(day.date), i]));

    s.real.forEach(t => {
        if (!t.date) return;
        const i = index.get(dayKey(t.date));
        if (i === undefined) return;
        visible[i][t.type === "income" ? "in" : "out"] += t.amount;
    });

    const max = Math.max(1, ...visible.map(d => Math.max(d.in, d.out)));

    const available = Math.max(280, chart.clientWidth || 600);
    const slot = Math.max(14, Math.min(40, available / visible.length));
    const width = Math.max(visible.length * slot + 36, available);
    const half = 90;
    const height = half * 2 + 26;

    let bars = "";
    let lastLabelX = -Infinity;

    visible.forEach((day, i) => {

        const x = i * slot;
        // Square-root scale keeps small days visible next to one big transfer
        const inH = Math.sqrt(day.in / max) * (half - 6);
        const outH = Math.sqrt(day.out / max) * (half - 6);
        const previous = visible[i - 1];
        const jump = previous && (day.date - previous.date) > 86400000 * 1.5;
        const wantLabel = i === 0 || jump || day.date.getDate() === 1 || (day.date.getDay() === 1 && slot >= 14 && i > 2);
        const showLabel = wantLabel && x - lastLabelX >= 56;

        if (showLabel) {
            lastLabelX = x;
        }

        bars += `
            ${jump ? `<line x1="${x}" x2="${x}" y1="6" y2="${half * 2 - 6}" stroke="var(--line)" stroke-dasharray="3 3"></line>` : ""}
            <g class="day" data-i="${i}" tabindex="0" role="button" aria-label="${formatDate(day.date)}: in ${formatNairaExact(day.in)}, out ${formatNairaExact(day.out)}">
                <rect class="hit" x="${x}" y="0" width="${slot}" height="${half * 2}" rx="4" fill="transparent"></rect>
                ${inH ? `<rect class="bar" x="${x + 3}" y="${half - inH}" width="${slot - 6}" height="${inH}" rx="3" fill="var(--in)"></rect>` : ""}
                ${outH ? `<rect class="bar" x="${x + 3}" y="${half}" width="${slot - 6}" height="${outH}" rx="3" fill="var(--out)"></rect>` : ""}
                ${showLabel ? `<text x="${x + 1}" y="${height - 4}" font-size="11" fill="var(--ink-3)">${day.date.toLocaleDateString("en-NG", { day: "numeric", month: "short" })}</text>` : ""}
            </g>
        `;
    });

    chart.innerHTML = `
        <svg viewBox="0 0 ${width} ${height}" width="${width}" height="${height}" role="img" aria-label="Money in and out per day">
            <line x1="0" x2="${width}" y1="${half}" y2="${half}" stroke="var(--line)" stroke-width="1"></line>
            ${bars}
        </svg>
    `;

    // Wide charts start scrolled to the latest days
    chart.scrollLeft = chart.scrollWidth;

    const pick = el => {
        chart.querySelectorAll(".day.is-picked").forEach(n => n.classList.remove("is-picked"));
        el.classList.add("is-picked");
        const day = visible[Number(el.dataset.i)];
        $("chartTip").innerHTML =
            `<strong>${formatDate(day.date)}</strong>: ` +
            `in <span style="color:var(--in)">${formatNairaExact(day.in)}</span>, ` +
            `out <span style="color:var(--out)">${formatNairaExact(day.out)}</span>`;
    };

    chart.querySelectorAll(".day").forEach(el => {
        el.addEventListener("click", () => pick(el));
        el.addEventListener("keydown", event => {
            if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                pick(el);
            }
        });
    });
}

// ---------- Budget ----------

function renderBudget(s) {

    const budget = readStore(BUDGET_KEY, null);
    const body = $("budgetBody");

    if (!s.end) {
        body.innerHTML = "";
        return;
    }

    // Latest month in the statements
    const month = s.end.getMonth();
    const year = s.end.getFullYear();

    const spent = s.expenseTx
        .filter(t => t.date && t.date.getMonth() === month && t.date.getFullYear() === year)
        .reduce((total, t) => total + t.amount, 0);

    const monthName = s.end.toLocaleDateString("en-NG", { month: "long", year: "numeric" });

    $("budgetEdit").textContent = budget ? "Change" : "Set budget";

    if (!budget) {
        body.innerHTML = `
            <p class="big-line">${formatNairaExact(spent)}</p>
            <p class="muted">spent in ${escapeHTML(monthName)}. Set a monthly limit to track it.</p>
        `;
        return;
    }

    const ratio = spent / budget;
    const left = budget - spent;

    body.innerHTML = `
        <p class="big-line">${formatNaira(spent)} <small>of ${formatNaira(budget)} in ${escapeHTML(monthName)}</small></p>
        <div class="meter ${ratio > 1 ? "is-over" : ratio > 0.8 ? "is-warn" : ""}">
            <span style="width:${Math.min(ratio, 1) * 100}%"></span>
        </div>
        <p class="muted">
            ${left >= 0
                ? `${formatNairaExact(left)} left to spend this month.`
                : `${formatNairaExact(Math.abs(left))} over your limit.`}
        </p>
    `;
}

$("budgetEdit").addEventListener("click", () => {
    const form = $("budgetForm");
    form.hidden = !form.hidden;
    if (!form.hidden) {
        $("budgetInput").value = readStore(BUDGET_KEY, "") || "";
        $("budgetInput").focus();
    }
});

$("budgetForm").addEventListener("submit", event => {
    event.preventDefault();
    const value = Number($("budgetInput").value);
    writeStore(BUDGET_KEY, value > 0 ? value : null);
    $("budgetForm").hidden = true;
    renderBudget(state.summary);
});

// ---------- Insights ----------

function renderInsights(s) {

    const items = [];
    const sumWhere = test => s.expenseTx.filter(test).reduce((total, t) => total + t.amount, 0);

    // Top category
    const top = Object.entries(s.categories).sort((a, b) => b[1] - a[1])[0];

    if (top && s.expenses) {
        items.push([
            `${top[0]} took ${Math.round((top[1] / s.expenses) * 100)}% of your spending`,
            `${formatNairaExact(top[1])} out of ${formatNairaExact(s.expenses)}.`
        ]);
    }

    // Over / under
    if (s.expenses > s.income && s.income > 0) {
        items.push([
            `You spent ${formatNaira(s.expenses - s.income)} more than came in`,
            "The difference came out of money you already had."
        ]);
    }

    // Airtime + data
    const telco = sumWhere(t => /airtime|mobile data|data plan|\d+(\.\d+)?\s*gb\b|recharge/i.test(t.description));

    if (telco > 0) {
        items.push([
            `${formatNaira(telco)} on airtime and data`,
            s.covered.length >= 28
                ? `About ${formatNaira(telco / (s.covered.length / 30.4))} a month.`
                : `Across ${s.covered.length} day${s.covered.length === 1 ? "" : "s"} of statements.`
        ]);
    }

    // Charges
    const feeTx = s.expenseTx.filter(t => /stamp duty|charge|\bfee\b|vat|withholding|levy|sms/i.test(t.description));

    if (feeTx.length) {
        items.push([
            `${formatNairaExact(feeTx.reduce((total, t) => total + t.amount, 0))} in bank charges`,
            `${feeTx.length} charge${feeTx.length === 1 ? "" : "s"} like stamp duty, SMS and transfer fees.`
        ]);
    }

    // Biggest single spend
    const biggest = [...s.expenseTx].sort((a, b) => b.amount - a.amount)[0];

    if (biggest) {
        items.push([
            `Biggest payment: ${formatNaira(biggest.amount)}`,
            `${biggest.description}${biggest.date ? `, ${formatDate(biggest.date)}` : ""}.`
        ]);
    }

    // Daily average
    if (s.covered.length && s.expenses) {
        const days = s.covered.length;
        items.push([
            `${formatNaira(s.expenses / days)} a day on average`,
            `Across ${days} day${days === 1 ? "" : "s"} of statements.`
        ]);
    }

    // Savings moves
    if (s.internal > 0) {
        items.push([
            `${formatNaira(s.internal)} moved between your own accounts`,
            "Savings top-ups and withdrawals aren't counted as income or spending."
        ]);
    }

    $("insights").innerHTML = items
        .slice(0, 6)
        .map(([title, text]) => `<li><strong>${escapeHTML(title)}</strong>${escapeHTML(text)}</li>`)
        .join("") || `<li>Add more statements to see patterns.</li>`;
}

// ---------- People ----------

function payeeName(t) {

    if (/airtime|mobile data|data plan|\d+(\.\d+)?\s*gb\b/i.test(t.description)) {
        return "Airtime and data";
    }

    if (t.category === "Bills" && /stamp duty|charge|fee|vat|levy|sms/i.test(t.description)) {
        return "Bank charges";
    }

    let name = t.description
        .replace(/^(transfer|trf|tfr)\s+(to|from)\s+/i, "")
        .replace(/^pos\s*transfer\s*-?\s*/i, "")
        .split(" · ")[0]
        .replace(/…$/, "")
        .trim();

    if (!name) {
        return "Unknown";
    }

    // "ABUBAKAR MUHAMMED" -> "Abubakar Muhammed"
    if (name === name.toUpperCase()) {
        name = name.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    }

    return name;
}

function displayName(t) {

    const direction = t.description.match(/^(?:transfer|trf|tfr)\s+(to|from)\s+/i);

    if (direction) {
        return (direction[1].toLowerCase() === "to" ? "To " : "From ") + payeeName(t);
    }

    let text = t.description;

    if (text === text.toUpperCase()) {
        text = text.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    }

    return text.charAt(0).toUpperCase() + text.slice(1);
}

function renderPeople(s) {

    const groups = new Map();

    s.expenseTx.forEach(t => {
        const name = payeeName(t);
        const group = groups.get(name) || { name, total: 0, count: 0 };
        group.total += t.amount;
        group.count += 1;
        groups.set(name, group);
    });

    const top = [...groups.values()].sort((a, b) => b.total - a.total).slice(0, 6);

    $("people").innerHTML = top.length
        ? top.map(p => `
            <li>
                <span class="avatar" aria-hidden="true">${escapeHTML(initials(p.name))}</span>
                <span class="who">
                    <strong>${escapeHTML(p.name)}</strong>
                    <span>${p.count} payment${p.count === 1 ? "" : "s"}</span>
                </span>
                <span class="amt" title="${formatNairaExact(p.total)}">${formatNaira(p.total)}</span>
            </li>
        `).join("")
        : `<li class="muted">No payments yet.</li>`;
}

function initials(name) {
    return name
        .split(/\s+/)
        .filter(Boolean)
        .slice(0, 2)
        .map(word => word[0])
        .join("")
        .toUpperCase();
}

// ---------- Transactions ----------

document.querySelectorAll(".segmented button").forEach(button => {
    button.addEventListener("click", () => {
        document.querySelectorAll(".segmented button").forEach(b => b.classList.remove("is-active"));
        button.classList.add("is-active");
        state.filter = button.dataset.filter;
        state.shown = TX_PREVIEW;
        renderTransactions();
    });
});

$("txSearch").addEventListener("input", event => {
    state.search = event.target.value.trim().toLowerCase();
    state.shown = TX_PREVIEW;
    renderTransactions();
});

$("txMore").addEventListener("click", () => {
    const expanded = state.shown > TX_PREVIEW;

    state.shown = expanded ? TX_PREVIEW : Infinity;
    renderTransactions();

    // Collapsing: jump back to the top of the panel
    if (expanded) {
        document.querySelector(".area-tx").scrollIntoView({ behavior: "smooth", block: "start" });
    }
});

function renderTransactions() {

    const s = state.summary;

    if (!s) return;

    let list = [...s.all].reverse();

    if (state.filter === "income") list = list.filter(t => !t.internal && t.type === "income");
    if (state.filter === "expense") list = list.filter(t => !t.internal && t.type === "expense");
    if (state.filter === "internal") list = list.filter(t => t.internal);
    if (state.filter === "all") list = list.filter(t => !t.internal);

    if (state.search) {
        list = list.filter(t =>
            (t.description + " " + t.category + " " + t.account)
                .toLowerCase()
                .includes(state.search)
        );
    }

    $("txMeta").textContent = `${list.length} transaction${list.length === 1 ? "" : "s"}`;

    if (!list.length) {
        $("txList").innerHTML = `<p class="tx-empty">No transactions match. Try another search or filter.</p>`;
        $("txMore").hidden = true;
        return;
    }

    const page = list.slice(0, state.shown);
    let html = "";
    let lastDay = "";

    page.forEach(t => {

        const day = t.date ? formatDate(t.date) : "No date";

        if (day !== lastDay) {

            const dayTotal = list
                .filter(x => (x.date ? formatDate(x.date) : "No date") === day && !x.internal)
                .reduce((total, x) => total + (x.type === "income" ? x.amount : -x.amount), 0);

            html += `<div class="tx-day"><span>${escapeHTML(day)}</span><span>${dayTotal ? (dayTotal > 0 ? "+" : "−") + formatNaira(Math.abs(dayTotal)) : ""}</span></div>`;
            lastDay = day;
        }

        const color = CATEGORY_COLORS[t.category] || "var(--c-other)";
        const amountClass = t.internal ? "is-internal" : t.type === "income" ? "is-in" : "";
        const sign = t.type === "income" ? "+" : "−";
        const time = t.date && (t.date.getHours() || t.date.getMinutes())
            ? t.date.toLocaleTimeString("en-NG", { hour: "numeric", minute: "2-digit" })
            : "";

        const title = displayName(t);

        html += `
            <div class="tx">
                <span class="dot" style="background:${color}"></span>
                <div class="tx-info">
                    <span class="tx-name" title="${escapeHTML(t.description)}">${escapeHTML(title)}</span>
                    <span class="tx-sub">${escapeHTML(t.category)}, ${escapeHTML(t.account)}${time ? ", " + time : ""}</span>
                </div>
                <span class="tx-amt ${amountClass}" title="${formatNairaExact(t.amount)}">${sign}${formatNairaExact(t.amount)}</span>
            </div>
        `;
    });

    $("txList").innerHTML = html;
    $("txMore").hidden = list.length <= TX_PREVIEW;
    $("txMore").textContent = state.shown > TX_PREVIEW
        ? "Show less"
        : `Show all ${list.length} transactions`;
    $("txMore").setAttribute("aria-expanded", String(state.shown > TX_PREVIEW));
}

// ---------- Goal ----------

function renderGoal(s) {

    const goal = readStore(GOAL_KEY, null);
    const body = $("goalBody");

    $("goalEdit").textContent = goal ? "Change" : "Set goal";

    if (!goal) {
        body.innerHTML = `<p class="muted">Pick something to save for and Fin-dash shows how close your balance is.</p>`;
        return;
    }

    const have = Math.max(0, s.balance === null ? s.income - s.expenses : s.balance);
    const ratio = Math.min(have / goal.target, 1);

    body.innerHTML = `
        <p class="big-line">${escapeHTML(goal.name)}</p>
        <div class="meter"><span style="width:${ratio * 100}%"></span></div>
        <p class="muted">
            ${formatNaira(have)} of ${formatNaira(goal.target)} (${Math.round(ratio * 100)}%).
            ${ratio >= 1 ? "You have enough." : `${formatNairaExact(goal.target - have)} to go.`}
        </p>
    `;
}

$("goalEdit").addEventListener("click", () => {
    const form = $("goalForm");
    form.hidden = !form.hidden;
    if (!form.hidden) {
        const goal = readStore(GOAL_KEY, null);
        $("goalName").value = goal ? goal.name : "";
        $("goalTarget").value = goal ? goal.target : "";
        $("goalName").focus();
    }
});

$("goalForm").addEventListener("submit", event => {
    event.preventDefault();
    const target = Number($("goalTarget").value);
    const name = $("goalName").value.trim();
    writeStore(GOAL_KEY, target > 0 && name ? { name, target } : null);
    $("goalForm").hidden = true;
    renderGoal(state.summary);
});

// ---------- Statements list ----------

function renderFiles() {

    $("files").innerHTML = state.statements.map(statement => {

        const dates = statement.transactions.map(t => t.date).filter(Boolean);
        const range = dates.length
            ? `${formatDate(dates[0])} to ${formatDate(dates[dates.length - 1])}`
            : "";

        return `
            <li>
                <div>
                    <strong>${escapeHTML(statement.bank)}</strong>
                    <span>${statement.transactions.length} transactions${range ? ", " + range : ""}</span>
                </div>
                <button type="button" data-remove="${statement.id}" aria-label="Remove ${escapeHTML(statement.name)}">Remove</button>
            </li>
        `;
    }).join("");

    $("files").querySelectorAll("[data-remove]").forEach(button => {
        button.addEventListener("click", () => {
            state.statements = state.statements.filter(s => s.id !== button.dataset.remove);
            saveStatements();
            render();
        });
    });
}

// ---------- Export + clear ----------

$("exportBtn").addEventListener("click", () => {

    const s = state.summary;

    if (!s) return;

    const rows = [["Date", "Account", "Description", "Category", "Type", "Amount", "Balance"]];

    s.all.forEach(t => {
        rows.push([
            t.date ? t.date.toISOString().slice(0, 10) : "",
            t.account,
            t.description,
            t.category,
            t.internal ? "savings move" : t.type === "income" ? "in" : "out",
            t.amount.toFixed(2),
            t.balance === null || t.balance === undefined ? "" : Number(t.balance).toFixed(2)
        ]);
    });

    const csv = rows
        .map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(","))
        .join("\n");

    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    link.download = "fin-dash-transactions.csv";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
});

$("clearBtn").addEventListener("click", () => {
    if (!confirm("Remove all statements from this device? Your budget and goal stay.")) return;
    state.statements = [];
    saveStatements();
    render();
});

// ========================================
// FORMATTING
// ========================================

function formatNaira(value) {

    const n = Math.abs(Number(value) || 0);
    const sign = value < 0 ? "−" : "";

    if (n >= 1e9) return `${sign}₦${trimDecimal(n / 1e9)}B`;
    if (n >= 1e6) return `${sign}₦${trimDecimal(n / 1e6)}M`;
    if (n >= 1e4) return `${sign}₦${trimDecimal(n / 1e3)}k`;

    return sign + "₦" + n.toLocaleString("en-NG", {
        minimumFractionDigits: n % 1 ? 2 : 0,
        maximumFractionDigits: 2
    });
}

function trimDecimal(value) {
    return value.toFixed(value >= 100 ? 0 : 1).replace(/\.0$/, "");
}

function formatNairaExact(value) {
    const n = Number(value) || 0;
    return (n < 0 ? "−" : "") + "₦" + Math.abs(n).toLocaleString("en-NG", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function escapeHTML(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

// ========================================
// GUILLOCHE (banknote engraving behind balance)
// ========================================

function drawGuilloche() {

    const svg = $("guilloche");
    const size = 600;
    const c = size / 2;
    let paths = "";

    // Layered hypotrochoid rosettes, like the line work on a naira note
    const layers = [
        { R: 210, r: 47, d: 64, turns: 47 },
        { R: 255, r: 59, d: 30, turns: 59 },
        { R: 288, r: 23, d: 10, turns: 23 }
    ];

    layers.forEach(({ R, r, d, turns }) => {

        let path = "";
        const steps = turns * 60;

        for (let i = 0; i <= steps; i++) {
            const t = (i / steps) * Math.PI * 2 * turns;
            const x = c + (R - r) * Math.cos(t) + d * Math.cos(((R - r) / r) * t);
            const y = c + (R - r) * Math.sin(t) - d * Math.sin(((R - r) / r) * t);
            path += (i ? "L" : "M") + x.toFixed(1) + " " + y.toFixed(1);
        }

        paths += `<path d="${path}"></path>`;
    });

    svg.setAttribute("viewBox", `0 0 ${size} ${size}`);
    svg.innerHTML = paths;
}

// ========================================
// LIBRARY LOADERS
// ========================================

function loadScript(src) {
    return new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.src = src;
        script.onload = resolve;
        script.onerror = () => reject(new Error("Couldn't load " + src));
        document.head.appendChild(script);
    });
}

async function loadXLSX() {
    if (window.XLSX) return;
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js");
}

async function loadPDFJS() {
    if (window.pdfjsLib) return;
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js");
    pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

// ========================================
// START
// ========================================

drawGuilloche();
loadStatements();
render();
