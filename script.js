// ========================================
// KUDA PDF PARSER
// ========================================

function parseKudaPDFRows(rows) {
    const transactions = [];
    let columns = null;

    for (const row of rows) {
        const text = row.items.map(i => i.text).join(" ").toLowerCase();

        if (text.includes("money in") && text.includes("money out")) {
            columns = detectKudaColumns(row);
            break;
        }
    }

    for (const row of rows) {
        const line = row.items.map(i => i.text).join(" ").replace(/\s+/g, " ").trim();
        const dateMatch = line.match(/\b\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}\b/);

        if (!dateMatch) continue;

        const amounts = extractAmountItems(row.items);
        if (!amounts.length) continue;

        let income = 0;
        let expense = 0;

        if (columns) {
            amounts.forEach(a => {
                const col = getClosestKudaColumn(a.x, columns);

                if (col === "moneyIn") income += a.value;
                if (col === "moneyOut") expense += a.value;
            });
        }

        // Kuda PDF fallback: transaction amount is usually before balance
        if (!income && !expense && amounts.length >= 2) {
            const amount = amounts[amounts.length - 2].value;

            if (looksLikeIncome(line)) income = amount;
            else expense = amount;
        }

        if (!income && !expense) continue;

        const description = cleanDescription(
            extractKudaDescription(row.items, dateMatch[0])
        );

        transactions.push({
            date: dateMatch[0],
            description,
            amount: income || expense,
            type: income ? "income" : "expense",
            category: categorize(description)
        });
    }

    return calculateFinancials(transactions);
}


// ========================================
// KUDA COLUMN DETECTION
// ========================================

function detectKudaColumns(row) {
    const c = {};

    row.items.forEach(item => {
        const t = item.text.toLowerCase().trim();

        if (t === "money in") c.moneyIn = item.x;
        if (t === "money out") c.moneyOut = item.x;
        if (t === "balance") c.balance = item.x;
    });

    return c.moneyIn !== undefined || c.moneyOut !== undefined ? c : null;
}

function getClosestKudaColumn(x, c) {
    const cols = [];

    if (c.moneyIn !== undefined)
        cols.push(["moneyIn", Math.abs(x - c.moneyIn)]);

    if (c.moneyOut !== undefined)
        cols.push(["moneyOut", Math.abs(x - c.moneyOut)]);

    if (!cols.length) return null;

    cols.sort((a, b) => a[1] - b[1]);

    return cols[0][0];
}


// ========================================
// CLEAN DESCRIPTION
// ========================================

function cleanDescription(value) {
    let text = String(value || "")
        .replace(/\s+/g, " ")
        .replace(/\b(reference|session|transaction id|payment reference|ref)\s*[:#-]?\s*/gi, "")
        .replace(/\b[A-Z0-9]{10,}\b/g, "")
        .replace(/\b[0-9a-f]{8}-[0-9a-f-]{20,}\b/gi, "")
        .replace(/[|]{2,}/g, " ")
        .replace(/\s+/g, " ")
        .trim();

    return text ? text.slice(0, 55) : "Bank transaction";
}


// ========================================
// COMPACT NAIRA
// ========================================

function formatNaira(value) {
    const n = Number(value || 0);
    const a = Math.abs(n);

    if (a >= 1000000)
        return "₦" + (n / 1000000).toFixed(2) + "M";

    if (a >= 1000)
        return "₦" + (n / 1000).toFixed(2) + "k";

    return "₦" + n.toLocaleString("en-NG", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function formatNairaExact(value) {
    return "₦" + Number(value || 0).toLocaleString("en-NG", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}
