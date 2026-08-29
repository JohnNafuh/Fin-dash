// ========================================
// FIN-DASH
// KUDA STATEMENT ANALYZER
// ========================================

const uploadButton = document.getElementById("chooseStatement");
const statementInput = document.getElementById("statementFile");

const uploadState = document.getElementById("statementUpload");
const analyzingState = document.getElementById("statementAnalyzing");
const completeState = document.getElementById("statementComplete");


// ========================================
// FILE UPLOAD
// ========================================

uploadButton?.addEventListener("click", () => {
    statementInput?.click();
});

statementInput?.addEventListener("change", () => {
    const file = statementInput.files?.[0];
    if (file) startAnalysis(file);
});

document.getElementById("analyzeAnother")?.addEventListener("click", () => {
    statementInput.value = "";

    uploadState.style.display = "block";
    analyzingState.style.display = "none";
    completeState.style.display = "none";
});


// ========================================
// ANALYSIS
// ========================================

async function startAnalysis(file) {

    uploadState.style.display = "none";
    completeState.style.display = "none";
    analyzingState.style.display = "flex";

    const fileName = document.getElementById("analyzingFileName");
    const progress = document.getElementById("analysisProgressBar");

    if (fileName) fileName.textContent = file.name;
    if (progress) progress.style.width = "0%";

    let value = 0;

    const loading = setInterval(() => {
        value += Math.random() * 8;
        value = Math.min(value, 90);

        if (progress) {
            progress.style.width = value + "%";
        }
    }, 250);

    try {

        const data = await analyzeStatement(file);

        clearInterval(loading);

        if (progress) progress.style.width = "100%";

        setTimeout(() => {
            updateDashboard(data);
            showComplete(data, file.name);
        }, 500);

    } catch (error) {

        clearInterval(loading);

        console.error("Fin-dash error:", error);
        showUploadError();
    }
}


// ========================================
// READ FILE
// ========================================

async function analyzeStatement(file) {

    const extension =
        file.name.split(".").pop().toLowerCase();


    if (extension === "csv") {
        return parseCSV(await file.text());
    }


    if (extension === "xlsx" || extension === "xls") {

        await loadXLSX();

        const workbook = XLSX.read(
            await file.arrayBuffer(),
            { type: "array" }
        );

        const sheet =
            workbook.Sheets[workbook.SheetNames[0]];

        const rows =
            XLSX.utils.sheet_to_json(
                sheet,
                {
                    header: 1,
                    defval: ""
                }
            );

        return parseRows(rows);
    }


    if (extension === "pdf") {

        await loadPDFJS();

        const pdf =
            await pdfjsLib.getDocument({
                data: await file.arrayBuffer()
            }).promise;

        const rows = [];

        for (let pageNo = 1; pageNo <= pdf.numPages; pageNo++) {

            const page = await pdf.getPage(pageNo);

            const content =
                await page.getTextContent();

            rows.push(
                ...groupPDFItems(content.items)
            );
        }

        return parseKudaPDFRows(rows);
    }


    throw new Error("Unsupported file type");
}


// ========================================
// CSV
// ========================================

function parseCSV(text) {

    const rows =
        text
            .split(/\r?\n/)
            .filter(Boolean)
            .map(line =>
                line
                    .split(",")
                    .map(cell =>
                        cell
                            .replace(/^"|"$/g, "")
                            .trim()
                    )
            );

    return parseRows(rows);
}


// ========================================
// CSV / XLSX
// ========================================

function parseRows(rows) {

    if (!rows.length) return emptyData();

    let headerIndex = -1;

    for (let i = 0; i < Math.min(rows.length, 20); i++) {

        const text =
            rows[i].join(" ").toLowerCase();

        if (
            text.includes("date") &&
            (
                text.includes("description") ||
                text.includes("money in") ||
                text.includes("money out") ||
                text.includes("amount")
            )
        ) {
            headerIndex = i;
            break;
        }
    }

    if (headerIndex < 0) headerIndex = 0;

    const headers =
        rows[headerIndex].map(v =>
            String(v).toLowerCase().trim()
        );

    const findColumn = (...names) =>
        headers.findIndex(header =>
            names.some(name =>
                header.includes(name)
            )
        );

    const dateIndex =
        findColumn("date", "transaction date");

    const descriptionIndex =
        findColumn(
            "description",
            "narration",
            "details",
            "merchant",
            "particular"
        );

    const moneyInIndex =
        findColumn(
            "money in",
            "credit",
            "deposit",
            "inflow"
        );

    const moneyOutIndex =
        findColumn(
            "money out",
            "debit",
            "withdrawal",
            "withdraw",
            "outflow"
        );

    const amountIndex =
        findColumn("amount", "value");

    const typeIndex =
        findColumn("type", "transaction type");

    const transactions = [];

    for (let i = headerIndex + 1; i < rows.length; i++) {

        const row = rows[i];

        if (!row?.length) continue;

        const description =
            cleanDescription(
                descriptionIndex >= 0
                    ? row[descriptionIndex]
                    : ""
            );

        const moneyIn =
            moneyInIndex >= 0
                ? parseMoney(row[moneyInIndex])
                : 0;

        const moneyOut =
            moneyOutIndex >= 0
                ? parseMoney(row[moneyOutIndex])
                : 0;

        const amount =
            amountIndex >= 0
                ? parseMoney(row[amountIndex])
                : 0;

        const type =
            typeIndex >= 0
                ? String(row[typeIndex] || "")
                : "";


        if (moneyIn > 0) {

            transactions.push({
                date: dateIndex >= 0 ? row[dateIndex] : "",
                description,
                amount: moneyIn,
                type: "income",
                category: categorize(description)
            });

            continue;
        }


        if (moneyOut > 0) {

            transactions.push({
                date: dateIndex >= 0 ? row[dateIndex] : "",
                description,
                amount: moneyOut,
                type: "expense",
                category: categorize(description)
            });

            continue;
        }


        if (amount !== 0) {

            const income =
                isIncome(type, description, amount);

            transactions.push({
                date: dateIndex >= 0 ? row[dateIndex] : "",
                description,
                amount: Math.abs(amount),
                type: income ? "income" : "expense",
                category: categorize(description)
            });
        }
    }

    return calculateFinancials(transactions);
}


// ========================================
// PDF ROW GROUPING
// ========================================

function groupPDFItems(items) {

    const rows = [];
    const tolerance = 4;

    items.forEach(item => {

        const text =
            String(item.str || "").trim();

        if (!text) return;

        const x = item.transform[4];
        const y = item.transform[5];

        let row =
            rows.find(r =>
                Math.abs(r.y - y) <= tolerance
            );

        if (!row) {
            row = {
                y,
                items: []
            };

            rows.push(row);
        }

        row.items.push({
            x,
            text
        });
    });

    return rows
        .sort((a, b) => b.y - a.y)
        .map(row => {
            row.items.sort((a, b) => a.x - b.x);
            return row;
        });
}


// ========================================
// KUDA PDF
// ========================================

function parseKudaPDFRows(rows) {

    const transactions = [];

    let columns = null;

    // Find Kuda headers first
    for (const row of rows) {

        const text =
            row.items
                .map(i => i.text)
                .join(" ")
                .toLowerCase();

        if (
            text.includes("money in") &&
            text.includes("money out") &&
            text.includes("balance")
        ) {

            columns = detectKudaColumns(row);
            break;
        }
    }


    if (!columns) {
        console.warn("Kuda columns not detected.");
        return emptyData();
    }


    for (const row of rows) {

        const line =
            row.items
                .map(i => i.text)
                .join(" ")
                .replace(/\s+/g, " ")
                .trim();

        if (!line) continue;

        const dateMatch =
            line.match(
                /\b\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}\b/
            );

        if (!dateMatch) continue;

        const date = dateMatch[0];

        const lower = line.toLowerCase();

        if (
            lower.includes("transaction date") ||
            lower.includes("money in") ||
            lower.includes("money out") ||
            lower.includes("opening balance") ||
            lower.includes("closing balance") ||
            lower.includes("account number") ||
            lower.includes("statement period") ||
            lower.includes("account name")
        ) {
            continue;
        }


        // --------------------------------
        // GET ONLY REAL MONEY VALUES
        // --------------------------------

        const amounts =
            extractAmountItems(row.items);

        if (!amounts.length) continue;


        let income = 0;
        let expense = 0;


        amounts.forEach(item => {

            const column =
                getKudaColumn(
                    item.x,
                    columns
                );

            if (column === "moneyIn") {
                income = item.value;
            }

            if (column === "moneyOut") {
                expense = item.value;
            }
        });


        // --------------------------------
        // FALLBACK
        // --------------------------------

        if (!income && !expense) {

            const transactionAmount =
                findTransactionAmount(
                    amounts,
                    columns
                );

            if (!transactionAmount) continue;

            if (looksLikeIncome(line)) {
                income = transactionAmount;
            } else {
                expense = transactionAmount;
            }
        }


        if (!income && !expense) continue;


        const description =
            cleanDescription(
                extractKudaDescription(
                    row.items,
                    date,
                    columns
                )
            );


        transactions.push({

            date,

            description:
                description || "Bank transaction",

            amount:
                income || expense,

            type:
                income ? "income" : "expense",

            category:
                categorize(
                    description
                )
        });
    }


    return calculateFinancials(
        transactions
    );
}


// ========================================
// KUDA COLUMNS
// ========================================

function detectKudaColumns(row) {

    const columns = {};

    row.items.forEach(item => {

        const text =
            item.text
                .toLowerCase()
                .trim();

        if (text === "date" || text.includes("transaction date")) {
            columns.date = item.x;
        }

        if (text === "description") {
            columns.description = item.x;
        }

        if (
            text === "money in" ||
            text.includes("money in")
        ) {
            columns.moneyIn = item.x;
        }

        if (
            text === "money out" ||
            text.includes("money out")
        ) {
            columns.moneyOut = item.x;
        }

        if (
            text === "balance" ||
            text.includes("balance")
        ) {
            columns.balance = item.x;
        }
    });

    return columns;
}


// ========================================
// GET KUDA COLUMN
// ========================================

function getKudaColumn(x, columns) {

    const candidates = [];

    if (columns.moneyIn !== undefined) {
        candidates.push({
            name: "moneyIn",
            distance: Math.abs(x - columns.moneyIn)
        });
    }

    if (columns.moneyOut !== undefined) {
        candidates.push({
            name: "moneyOut",
            distance: Math.abs(x - columns.moneyOut)
        });
    }

    if (columns.balance !== undefined) {
        candidates.push({
            name: "balance",
            distance: Math.abs(x - columns.balance)
        });
    }

    candidates.sort(
        (a, b) => a.distance - b.distance
    );

    if (!candidates.length) return null;

    const closest = candidates[0];

    // Important:
    // Never allow Balance to become income/expense.
    if (closest.name === "balance") {
        return null;
    }

    // Prevent unrelated numbers from being classified.
    if (closest.distance > 70) {
        return null;
    }

    return closest.name;
}


// ========================================
// AMOUNT EXTRACTION
// ========================================

function extractAmountItems(items) {

    const results = [];

    items.forEach(item => {

        const text =
            String(item.text || "").trim();

        // Kuda transaction amounts normally
        // contain decimals or currency formatting.
        if (
            !/(?:₦|NGN|\d[\d,]*\.\d{2})/.test(text)
        ) {
            return;
        }

        const value = parseMoney(text);

        if (!value) return;

        results.push({
            x: item.x,
            text,
            value: Math.abs(value)
        });
    });

    return results;
}


// ========================================
// FIND TRANSACTION AMOUNT
// ========================================

function findTransactionAmount(amounts, columns) {

    // Ignore balance completely.
    const candidates =
        amounts.filter(item => {

            if (columns.balance === undefined) {
                return true;
            }

            return (
                Math.abs(
                    item.x - columns.balance
                ) > 70
            );
        });


    if (!candidates.length) {
        return 0;
    }


    // Prefer the amount closest to Money In
    // or Money Out rather than arbitrary numbers.
    const transactionColumns = [];

    if (columns.moneyIn !== undefined) {
        transactionColumns.push(columns.moneyIn);
    }

    if (columns.moneyOut !== undefined) {
        transactionColumns.push(columns.moneyOut);
    }


    if (transactionColumns.length) {

        candidates.sort((a, b) => {

            const distanceA =
                Math.min(
                    ...transactionColumns.map(
                        x => Math.abs(a.x - x)
                    )
                );

            const distanceB =
                Math.min(
                    ...transactionColumns.map(
                        x => Math.abs(b.x - x)
                    )
                );

            return distanceA - distanceB;
        });
    }


    return candidates[0]?.value || 0;
}


// ========================================
// DESCRIPTION
// ========================================

function extractKudaDescription(
    items,
    date,
    columns
) {

    const parts = [];

    items.forEach(item => {

        const text =
            String(item.text || "").trim();

        if (!text) return;

        if (
            text === date ||
            text.includes(date)
        ) {
            return;
        }


        // Ignore anything sitting in
        // Money In, Money Out or Balance.
        if (columns.moneyIn !== undefined &&
            Math.abs(item.x - columns.moneyIn) < 70) {
            return;
        }

        if (columns.moneyOut !== undefined &&
            Math.abs(item.x - columns.moneyOut) < 70) {
            return;
        }

        if (columns.balance !== undefined &&
            Math.abs(item.x - columns.balance) < 70) {
            return;
        }


        // Ignore numeric values.
        if (
            /(?:₦|NGN|\d[\d,]*\.\d{2})/.test(text)
        ) {
            return;
        }


        const lower = text.toLowerCase();

        if (
            lower === "date" ||
            lower === "description" ||
            lower === "money in" ||
            lower === "money out" ||
            lower === "balance"
        ) {
            return;
        }

        parts.push(text);
    });


    return parts.join(" ");
}


// ========================================
// CLEAN DESCRIPTION
// ========================================

function cleanDescription(value) {

    let text =
        String(value || "")
            .replace(/\s+/g, " ")
            .trim();


    // Remove UUIDs
    text =
        text.replace(
            /\b[0-9a-f]{8}-[0-9a-f-]{20,}\b/gi,
            ""
        );


    // Remove long references / IDs
    text =
        text.replace(
            /\b[A-Z0-9]{10,}\b/g,
            ""
        );


    text =
        text.replace(
            /\b(reference|session|transaction id|payment reference|ref)\s*[:#-]?\s*/gi,
            ""
        );


    text =
        text.replace(
            /[|]{2,}/g,
            " "
        );


    text =
        text.replace(/\s+/g, " ").trim();


    if (!text) {
        return "Bank transaction";
    }


    // Keep recent transactions compact.
    if (text.length > 42) {
        text =
            text.slice(0, 42).trim() + "…";
    }


    return text;
}


// ========================================
// INCOME DETECTION
// ========================================

function isIncome(type, description, amount) {

    const text =
        `${type} ${description}`.toLowerCase();

    return (
        /credit|deposit|salary|income|received|refund|cashback|money in|inflow/.test(text) ||
        Number(amount) < 0
    );
}


function looksLikeIncome(line) {

    return /credit|deposit|salary|received|refund|cashback|money in|inflow|incoming|funded/i.test(
        line
    );
}


// ========================================
// CATEGORIES
// ========================================

function categorize(description) {

    const text =
        String(description).toLowerCase();


    if (
        /food|restaurant|eat|chicken|pizza|grocery|market|supermarket|shoprite|foodco|meal|kitchen|cafe|bakery/.test(text)
    ) {
        return "Food";
    }


    if (
        /uber|bolt|taxi|transport|fuel|petrol|gas|bus|car|ride|indrive|shell|total|mobil/.test(text)
    ) {
        return "Transport";
    }


    if (
        /electric|ikeja|aedc|phcn|water|internet|airtel|mtn|glo|9mobile|dstv|gotv|bill|utility|data|recharge|subscription/.test(text)
    ) {
        return "Bills";
    }


    if (
        /netflix|spotify|showmax|movie|cinema|game|entertainment|club|concert|music|apple music/.test(text)
    ) {
        return "Entertainment";
    }


    if (
        /shop|store|amazon|jumia|konga|purchase|pos|mall|fashion|clothing/.test(text)
    ) {
        return "Shopping";
    }


    if (
        /transfer|bank transfer|send money|sent to|trf|tfr|beneficiary/.test(text)
    ) {
        return "Transfers";
    }


    return "Other";
}


// ========================================
// FINANCIAL CALCULATIONS
// ========================================

function calculateFinancials(transactions) {

    let income = 0;
    let expenses = 0;

    const categories = {
        Food: 0,
        Transport: 0,
        Bills: 0,
        Entertainment: 0,
        Shopping: 0,
        Transfers: 0,
        Other: 0
    };


    transactions.forEach(transaction => {

        if (transaction.type === "income") {

            income += transaction.amount;

        } else {

            expenses += transaction.amount;

            if (
                categories[transaction.category] !== undefined
            ) {
                categories[transaction.category] +=
                    transaction.amount;
            }
        }
    });


    return {
        transactions,
        income,
        expenses,
        savings: income - expenses,
        balance: income - expenses,
        categories
    };
}


function emptyData() {

    return {
        transactions: [],
        income: 0,
        expenses: 0,
        savings: 0,
        balance: 0,

        categories: {
            Food: 0,
            Transport: 0,
            Bills: 0,
            Entertainment: 0,
            Shopping: 0,
            Transfers: 0,
            Other: 0
        }
    };
}


// ========================================
// MONEY
// ========================================

function parseMoney(value) {

    if (value === undefined || value === null) {
        return 0;
    }

    const raw = String(value).trim();

    if (!raw) return 0;

    const negative =
        raw.includes("-") ||
        (raw.includes("(") && raw.includes(")"));

    const number =
        parseFloat(
            raw.replace(
                /[₦$£NGN,\s()]/gi,
                ""
            )
        );

    if (isNaN(number)) return 0;

    return negative
        ? -Math.abs(number)
        : number;
}


// ========================================
// COMPACT FORMAT
// ========================================

function formatNaira(value) {

    const number =
        Number(value || 0);

    const sign =
        number < 0 ? "-" : "";

    const absolute =
        Math.abs(number);

    let result;


    if (absolute >= 1000000000) {

        result =
            (absolute / 1000000000)
                .toFixed(2) + "B";

    } else if (absolute >= 1000000) {

        result =
            (absolute / 1000000)
                .toFixed(2) + "M";

    } else if (absolute >= 1000) {

        result =
            (absolute / 1000)
                .toFixed(1) + "k";

    } else {

        result =
            absolute.toLocaleString(
                "en-NG",
                {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2
                }
            );
    }


    return sign + "₦" + result;
}


// ========================================
// EXACT FORMAT
// ========================================

function formatNairaExact(value) {

    return (
        Number(value || 0)
            .toLocaleString(
                "en-NG",
                {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2
                }
            )
            .replace(
                /^-/,
                "-₦"
            )
            .replace(
                /^/,
                Number(value || 0) < 0
                    ? ""
                    : "₦"
            )
    );
}


// ========================================
// DASHBOARD
// ========================================

function updateDashboard(data) {

    const balance =
        document.querySelector(".balance h2");

    const cards =
        document.querySelectorAll(
            ".summary .card h3"
        );


    if (balance) {

        balance.textContent =
            formatNaira(data.balance);

        balance.title =
            formatNairaExact(data.balance);
    }


    if (cards.length >= 3) {

        cards[0].textContent =
            formatNaira(data.income);

        cards[0].title =
            formatNairaExact(data.income);


        cards[1].textContent =
            formatNaira(data.expenses);

        cards[1].title =
            formatNairaExact(data.expenses);


        cards[2].textContent =
            formatNaira(data.savings);

        cards[2].title =
            formatNairaExact(data.savings);
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

    if (!list) return;


    if (!data.transactions.length) {

        list.innerHTML = `
            <div class="transaction empty-transaction">
                <p>No transactions found.</p>
            </div>
        `;

        return;
    }


    list.innerHTML =
        data.transactions
            .slice(-8)
            .reverse()
            .map(transaction => {

                const exact =
                    formatNairaExact(
                        transaction.amount
                    );

                const sign =
                    transaction.type === "income"
                        ? "+"
                        : "-";

                return `
                    <div class="transaction">

                        <div>
                            <strong>
                                ${escapeHTML(
                                    transaction.description
                                )}
                            </strong>

                            <small>
                                ${escapeHTML(
                                    transaction.category
                                )}
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
                            ${sign}${formatNaira(
                                transaction.amount
                            )}
                        </strong>

                    </div>
                `;
            })
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

    if (!items.length) return;


    const categories = [
        "Food",
        "Transport",
        "Bills",
        "Entertainment"
    ];


    const values =
        categories.map(
            category =>
                data.categories[category] || 0
        );


    const max =
        Math.max(...values, 1);


    items.forEach((item, index) => {

        const category =
            categories[index];

        if (!category) return;


        const value =
            data.categories[category] || 0;


        const amount =
            item.querySelector("span");


        const bar =
            item.querySelector(
                ".progress-bar"
            );


        if (amount) {

            amount.textContent =
                formatNaira(value);

            amount.title =
                formatNairaExact(value);
        }


        if (bar) {

            const percentage =
                value > 0
                    ? (value / max) * 100
                    : 0;

            bar.style.width =
                Math.min(
                    percentage,
                    100
                ) + "%";
        }
    });
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
                (spent / budget) * 100,
                100
            )
            : 0;


    if (amount) {

        amount.textContent =
            formatNaira(spent);

        amount.title =
            formatNairaExact(spent);
    }


    if (description) {

        description.textContent =
            `of ${formatNaira(budget)} available`;

        description.title =
            formatNairaExact(budget);
    }


    if (bar) {
        bar.style.width =
            percentage + "%";
    }


    if (message) {

        if (!budget) {

            message.textContent =
                "No income was detected.";

        } else if (spent > budget) {

            message.textContent =
                "Spending is higher than recorded income.";

        } else {

            message.textContent =
                `${Math.round(
                    percentage
                )}% of available income spent.`;
        }
    }
}


// ========================================
// COMPLETE
// ========================================

function showComplete(data, fileName) {

    analyzingState.style.display = "none";
    completeState.style.display = "flex";


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


    if (name) name.textContent = fileName;

    if (count) {
        count.textContent =
            data.transactions.length;
    }

    if (income) {

        income.textContent =
            formatNaira(data.income);

        income.title =
            formatNairaExact(data.income);
    }

    if (expenses) {

        expenses.textContent =
            formatNaira(data.expenses);

        expenses.title =
            formatNairaExact(data.expenses);
    }
}


// ========================================
// ERROR
// ========================================

function showUploadError() {

    analyzingState.style.display = "none";
    completeState.style.display = "none";
    uploadState.style.display = "block";


    const title =
        uploadState.querySelector("h2");

    const message =
        uploadState.querySelector("p");


    if (title) {
        title.textContent =
            "Couldn't analyze statement";
    }

    if (message) {
        message.textContent =
            "Make sure you're using a valid Kuda CSV, XLSX, XLS or PDF statement.";
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
// SECURITY
// ========================================

function escapeHTML(value) {

    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}


// ========================================
// LOAD XLSX
// ========================================

function loadXLSX() {

    return new Promise((resolve, reject) => {

        if (window.XLSX) {
            resolve();
            return;
        }


        const script =
            document.createElement("script");

        script.src =
            "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";

        script.onload = resolve;
        script.onerror = reject;

        document.head.appendChild(script);
    });
}


// ========================================
// LOAD PDF.JS
// ========================================

function loadPDFJS() {

    return new Promise((resolve, reject) => {

        if (window.pdfjsLib) {
            resolve();
            return;
        }


        const script =
            document.createElement("script");

        script.src =
            "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";

        script.onload = () => {

            if (!window.pdfjsLib) {
                reject(
                    new Error("PDF.js failed")
                );
                return;
            }

            pdfjsLib
                .GlobalWorkerOptions
                .workerSrc =
                "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

            resolve();
        };

        script.onerror = reject;

        document.head.appendChild(script);
    });
}
