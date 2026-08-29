// ========================================
// FIN-DASH
// KUDA STATEMENT ANALYZER + LIVE MARKET
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

    const extension =
        file.name
            .split(".")
            .pop()
            .toLowerCase();

    // CSV
    if (extension === "csv") {

        const text = await file.text();

        return parseCSV(text);
    }

    // XLSX / XLS
    if (
        extension === "xlsx" ||
        extension === "xls"
    ) {

        await loadXLSX();

        const buffer =
            await file.arrayBuffer();

        const workbook =
            XLSX.read(
                buffer,
                {
                    type: "array"
                }
            );

        const sheet =
            workbook.Sheets[
                workbook.SheetNames[0]
            ];

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

    // PDF
    if (extension === "pdf") {

        await loadPDFJS();

        const buffer =
            await file.arrayBuffer();

        const pdf =
            await pdfjsLib.getDocument({
                data: buffer
            }).promise;

        let rows = [];

        for (
            let pageNo = 1;
            pageNo <= pdf.numPages;
            pageNo++
        ) {

            const page =
                await pdf.getPage(pageNo);

            const content =
                await page.getTextContent();

            const pageRows =
                groupPDFItems(
                    content.items
                );

            rows.push(...pageRows);
        }

        return parseKudaPDFRows(rows);
    }

    throw new Error("Unsupported file type");
}

// ========================================
// CSV PARSER
// ========================================

function parseCSV(text) {

    const rows =
        text
            .split(/\r?\n/)
            .filter(line => line.trim())
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
// XLSX / CSV ROW PARSER
// ========================================

function parseRows(rows) {

    if (!rows.length) {
        return emptyData();
    }

    let headerIndex = -1;

    for (
        let i = 0;
        i < Math.min(rows.length, 20);
        i++
    ) {

        const text =
            rows[i]
                .join(" ")
                .toLowerCase();

        if (
            text.includes("date") &&
            (
                text.includes("description") ||
                text.includes("narration") ||
                text.includes("amount") ||
                text.includes("money out") ||
                text.includes("money in")
            )
        ) {

            headerIndex = i;
            break;
        }
    }

    if (headerIndex < 0) {
        headerIndex = 0;
    }

    const headers =
        rows[headerIndex]
            .map(value =>
                String(value)
                    .toLowerCase()
                    .trim()
            );

    const findColumn = (...names) =>
        headers.findIndex(header =>
            names.some(name =>
                header.includes(name)
            )
        );

    const dateIndex =
        findColumn(
            "date",
            "transaction date"
        );

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
        findColumn(
            "amount",
            "value"
        );

    const typeIndex =
        findColumn(
            "type",
            "transaction type"
        );

    const transactions = [];

    for (
        let i = headerIndex + 1;
        i < rows.length;
        i++
    ) {

        const row = rows[i];

        if (!row || !row.length) {
            continue;
        }

        const description =
            descriptionIndex >= 0
                ? String(
                    row[descriptionIndex] || ""
                )
                : "";

        const moneyIn =
            moneyInIndex >= 0
                ? parseMoney(
                    row[moneyInIndex]
                )
                : 0;

        const moneyOut =
            moneyOutIndex >= 0
                ? parseMoney(
                    row[moneyOutIndex]
                )
                : 0;

        const amount =
            amountIndex >= 0
                ? parseMoney(
                    row[amountIndex]
                )
                : 0;

        const type =
            typeIndex >= 0
                ? String(
                    row[typeIndex] || ""
                )
                : "";

        if (moneyIn > 0) {

            transactions.push({

                date:
                    dateIndex >= 0
                        ? row[dateIndex]
                        : "",

                description:
                    cleanDescription(
                        description
                    ),

                amount:
                    moneyIn,

                type:
                    "income",

                category:
                    categorize(
                        description
                    )

            });

            continue;
        }

        if (moneyOut > 0) {

            transactions.push({

                date:
                    dateIndex >= 0
                        ? row[dateIndex]
                        : "",

                description:
                    cleanDescription(
                        description
                    ),

                amount:
                    moneyOut,

                type:
                    "expense",

                category:
                    categorize(
                        description
                    )

            });

            continue;
        }

        if (amount !== 0) {

            const income =
                isIncome(
                    type,
                    description,
                    amount
                );

            transactions.push({

                date:
                    dateIndex >= 0
                        ? row[dateIndex]
                        : "",

                description:
                    cleanDescription(
                        description
                    ),

                amount:
                    Math.abs(amount),

                type:
                    income
                        ? "income"
                        : "expense",

                category:
                    categorize(
                        description
                    )

            });

        }

    }

    return calculateFinancials(
        transactions
    );
}

// ========================================
// PDF ITEM GROUPING
// ========================================

function groupPDFItems(items) {

    const rows = [];

    const tolerance = 4;

    items.forEach(item => {

        const text =
            String(
                item.str || ""
            ).trim();

        if (!text) {
            return;
        }

        const x =
            item.transform[4];

        const y =
            item.transform[5];

        let row =
            rows.find(existing =>
                Math.abs(
                    existing.y - y
                ) <= tolerance
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
        .sort(
            (a, b) =>
                b.y - a.y
        );
}

// ========================================
// KUDA PDF PARSER
// ========================================

function parseKudaPDFRows(rows) {

    const transactions = [];

    let columnPositions = null;

    // ------------------------------------
    // FIND KUDA COLUMN HEADERS
    // ------------------------------------

    for (const row of rows) {

        const sortedItems =
            [...row.items].sort(
                (a, b) => a.x - b.x
            );

        const fullText =
            sortedItems
                .map(item => item.text)
                .join(" ")
                .toLowerCase()
                .replace(/\s+/g, " ")
                .trim();

        if (
            fullText.includes("money in") &&
            fullText.includes("money out")
        ) {

            columnPositions =
                detectKudaColumns(
                    row
                );

            break;
        }
    }

    // ------------------------------------
    // READ EACH PDF ROW
    // ------------------------------------

    for (const row of rows) {

        const sortedItems =
            [...row.items].sort(
                (a, b) => a.x - b.x
            );

        const line =
            sortedItems
                .map(item => item.text)
                .join(" ")
                .replace(/\s+/g, " ")
                .trim();

        if (!line) {
            continue;
        }

        const lower =
            line.toLowerCase();

        // Ignore headers / account information
        if (
            lower.includes("transaction date") ||
            lower.includes("money in") ||
            lower.includes("money out") ||
            lower.includes("opening balance") ||
            lower.includes("closing balance") ||
            lower.includes("account number") ||
            lower.includes("statement period") ||
            lower.includes("account name") ||
            lower === "date description"
        ) {

            continue;
        }

        // --------------------------------
        // FIND DATE
        // --------------------------------

        const dateMatch =
            line.match(
                /\b\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}\b/
            );

        if (!dateMatch) {
            continue;
        }

        const date =
            dateMatch[0];

        // --------------------------------
        // FIND AMOUNTS WITH POSITIONS
        // --------------------------------

        const amountItems =
            extractAmountItems(
                sortedItems
            );

        if (!amountItems.length) {
            continue;
        }

        let incomeAmount = 0;
        let expenseAmount = 0;

        // --------------------------------
        // POSITION-BASED KUDA PARSING
        // --------------------------------

        if (columnPositions) {

            const incomeCandidates = [];
            const expenseCandidates = [];

            amountItems.forEach(item => {

                const closest =
                    getKudaAmountColumn(
                        item.x,
                        columnPositions
                    );

                if (
                    closest === "moneyIn"
                ) {

                    incomeCandidates.push(
                        item.value
                    );

                }

                if (
                    closest === "moneyOut"
                ) {

                    expenseCandidates.push(
                        item.value
                    );

                }

            });

            if (
                incomeCandidates.length
            ) {

                incomeAmount =
                    Math.max(
                        ...incomeCandidates
                    );

            }

            if (
                expenseCandidates.length
            ) {

                expenseAmount =
                    Math.max(
                        ...expenseCandidates
                    );

            }

        }

        // --------------------------------
        // FALLBACK PARSING
        // --------------------------------

        if (
            incomeAmount === 0 &&
            expenseAmount === 0
        ) {

            const parsed =
                amountItems
                    .map(item =>
                        item.value
                    )
                    .filter(
                        value =>
                            value !== 0
                    );

            if (!parsed.length) {
                continue;
            }

            /*
             * Kuda statements normally contain:
             *
             * Transaction amount
             * Balance
             *
             * Therefore use the transaction
             * amount immediately before balance.
             */

            if (parsed.length >= 2) {

                const transactionAmount =
                    Math.abs(
                        parsed[
                            parsed.length - 2
                        ]
                    );

                if (
                    transactionAmount > 0
                ) {

                    if (
                        looksLikeIncome(
                            line
                        )
                    ) {

                        incomeAmount =
                            transactionAmount;

                    } else {

                        expenseAmount =
                            transactionAmount;

                    }

                }

            } else {

                const transactionAmount =
                    Math.abs(
                        parsed[0]
                    );

                if (
                    looksLikeIncome(
                        line
                    )
                ) {

                    incomeAmount =
                        transactionAmount;

                } else {

                    expenseAmount =
                        transactionAmount;

                }

            }

        }

        // --------------------------------
        // SKIP NON-TRANSACTION ROWS
        // --------------------------------

        if (
            incomeAmount === 0 &&
            expenseAmount === 0
        ) {

            continue;
        }

        // --------------------------------
        // CLEAN DESCRIPTION
        // --------------------------------

        let description =
            extractKudaDescription(
                sortedItems,
                date
            );

        description =
            cleanDescription(
                description
            );

        if (!description) {
            description =
                "Bank transaction";
        }

        // --------------------------------
        // ADD TRANSACTION
        // --------------------------------

        if (incomeAmount > 0) {

            transactions.push({

                date,

                description,

                amount:
                    incomeAmount,

                type:
                    "income",

                category:
                    categorize(
                        description
                    )

            });

        } else if (
            expenseAmount > 0
        ) {

            transactions.push({

                date,

                description,

                amount:
                    expenseAmount,

                type:
                    "expense",

                category:
                    categorize(
                        description
                    )

            });

        }

    }

    return calculateFinancials(
        transactions
    );
}

// ========================================
// DETECT KUDA COLUMNS
// ========================================

function detectKudaColumns(row) {

    const items =
        [...row.items].sort(
            (a, b) => a.x - b.x
        );

    const columns = {};

    items.forEach(item => {

        const text =
            String(item.text || "")
                .toLowerCase()
                .trim();

        if (
            text === "balance" ||
            text.includes("balance")
        ) {

            columns.balance =
                item.x;

        }

        if (
            text === "description" ||
            text.includes("description")
        ) {

            columns.description =
                item.x;

        }

        if (
            text === "date" ||
            text.includes("transaction date")
        ) {

            columns.date =
                item.x;

        }

        if (
            text === "money in"
        ) {

            columns.moneyIn =
                item.x;

        }

        if (
            text === "money out"
        ) {

            columns.moneyOut =
                item.x;

        }

    });

    // Handle PDF text where "Money" and "In"
    // or "Money" and "Out" are separate items.

    for (
        let i = 0;
        i < items.length - 1;
        i++
    ) {

        const first =
            String(
                items[i].text || ""
            )
                .toLowerCase()
                .trim();

        const second =
            String(
                items[i + 1].text || ""
            )
                .toLowerCase()
                .trim();

        if (
            first === "money" &&
            second === "in"
        ) {

            columns.moneyIn =
                (
                    items[i].x +
                    items[i + 1].x
                ) / 2;

        }

        if (
            first === "money" &&
            second === "out"
        ) {

            columns.moneyOut =
                (
                    items[i].x +
                    items[i + 1].x
                ) / 2;

        }

    }

    return (
        columns.moneyIn !== undefined ||
        columns.moneyOut !== undefined
    )
        ? columns
        : null;
}

// ========================================
// GET KUDA AMOUNT COLUMN
// ========================================

function getKudaAmountColumn(
    x,
    columns
) {

    const candidates = [];

    if (
        columns.moneyIn !== undefined
    ) {

        candidates.push({

            name:
                "moneyIn",

            distance:
                Math.abs(
                    x - columns.moneyIn
                )

        });

    }

    if (
        columns.moneyOut !== undefined
    ) {

        candidates.push({

            name:
                "moneyOut",

            distance:
                Math.abs(
                    x - columns.moneyOut
                )

        });

    }

    if (!candidates.length) {
        return null;
    }

    candidates.sort(
        (a, b) =>
            a.distance - b.distance
    );

    return candidates[0].name;
}

// ========================================
// GET CLOSEST KUDA COLUMN
// ========================================

function getClosestKudaColumn(
    x,
    columns
) {

    return getKudaAmountColumn(
        x,
        columns
    );
}

// ========================================
// EXTRACT AMOUNTS FROM PDF ITEMS
// ========================================

function extractAmountItems(items) {

    const results = [];

    items.forEach(item => {

        const text =
            String(
                item.text || ""
            )
                .trim();

        if (!text) {
            return;
        }

        const looksFinancial =
            /^(?:₦|NGN|\$|£)?\s*-?\s*\d[\d,]*(?:\.\d{1,2})?$/.test(
                text
            );

        if (!looksFinancial) {
            return;
        }

        const value =
            parseMoney(text);

        if (
            Number.isFinite(value) &&
            value !== 0
        ) {

            results.push({

                x:
                    item.x,

                text,

                value:
                    Math.abs(value)

            });

        }

    });

    return results;
}

// ========================================
// EXTRACT KUDA DESCRIPTION
// ========================================

function extractKudaDescription(
    items,
    date
) {

    const descriptionParts = [];

    items.forEach(item => {

        const text =
            String(
                item.text || ""
            ).trim();

        if (!text) {
            return;
        }

        // Remove date
        if (
            text === date ||
            text.includes(date)
        ) {

            const remaining =
                text
                    .replace(
                        date,
                        ""
                    )
                    .trim();

            if (
                remaining
            ) {

                descriptionParts.push(
                    remaining
                );

            }

            return;
        }

        // Ignore amounts
        if (
            /^(?:₦|NGN|\$|£)?\s*-?\s*\d[\d,]*(?:\.\d{1,2})?$/.test(
                text
            )
        ) {

            return;
        }

        // Ignore obvious headers
        const lower =
            text.toLowerCase();

        if (
            lower === "date" ||
            lower === "description" ||
            lower === "money in" ||
            lower === "money out" ||
            lower === "balance" ||
            lower === "money"
        ) {

            return;
        }

        descriptionParts.push(
            text
        );

    });

    return descriptionParts
        .join(" ")
        .replace(/\s+/g, " ")
        .trim();
}

// ========================================
// INCOME DETECTION
// ========================================

function isIncome(
    type,
    description,
    amount
) {

    const text =
        (
            type +
            " " +
            description
        )
            .toLowerCase();

    if (
        text.includes("credit") ||
        text.includes("deposit") ||
        text.includes("salary") ||
        text.includes("income") ||
        text.includes("received") ||
        text.includes("refund") ||
        text.includes("cashback") ||
        text.includes("money in") ||
        text.includes("inflow")
    ) {

        return true;
    }

    return Number(amount) < 0;
}

// ========================================
// KUDA INCOME HEURISTIC
// ========================================

function looksLikeIncome(line) {

    const text =
        line.toLowerCase();

    return (
        /credit|deposit|salary|received|refund|cashback|money in|inflow|incoming|funded/.test(
            text
        )
    );
}

// ========================================
// CATEGORY
// ========================================

function categorize(description) {

    const text =
        String(description || "")
            .toLowerCase();

    // FOOD
    if (
        /food|restaurant|eat|chicken|pizza|grocery|market|supermarket|shoprite|foodco|meal|kitchen|cafe|bakery|mcdonald|domino|kfc/.test(
            text
        )
    ) {

        return "Food";
    }

    // TRANSPORT
    if (
        /uber|bolt|taxi|transport|fuel|petrol|gas|bus|car|ride|indrive|shell|total|mobil/.test(
            text
        )
    ) {

        return "Transport";
    }

    // BILLS
    if (
        /electric|ikeja|aedc|phcn|water|internet|airtel|mtn|glo|9mobile|dstv|gotv|bill|utility|data|recharge|subscription/.test(
            text
        )
    ) {

        return "Bills";
    }

    // ENTERTAINMENT
    if (
        /netflix|spotify|showmax|movie|cinema|game|entertainment|club|concert|music|apple music/.test(
            text
        )
    ) {

        return "Entertainment";
    }

    // SHOPPING
    if (
        /shop|store|amazon|jumia|konga|purchase|pos|mall|fashion|clothing|ikeja city mall/.test(
            text
        )
    ) {

        return "Shopping";
    }

    // TRANSFERS
    if (
        /transfer|bank transfer|send money|sent to|trf|tfr|beneficiary/.test(
            text
        )
    ) {

        return "Transfers";
    }

    return "Other";
}

// ========================================
// CLEAN DESCRIPTION
// ========================================

function cleanDescription(value) {

    let text =
        String(value || "")
            .replace(/\s+/g, " ")
            .trim();

    // Remove dates
    text =
        text.replace(
            /\b\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}\b/g,
            ""
        );

    // Remove common Kuda reference labels
    text =
        text.replace(
            /\b(reference|session|transaction id|payment reference|ref|transaction reference)\s*[:#-]?\s*/gi,
            ""
        );

    // Remove UUID-like references
    text =
        text.replace(
            /\b[0-9a-f]{8}-[0-9a-f-]{20,}\b/gi,
            ""
        );

    // Remove long transaction/reference numbers
    text =
        text.replace(
            /\b\d{10,}\b/g,
            ""
        );

    // Remove long alphanumeric IDs
    text =
        text.replace(
            /\b[A-Z0-9]{14,}\b/g,
            ""
        );

    // Remove currency values
    text =
        text.replace(
            /(?:₦|NGN)\s*[\d,]+(?:\.\d{1,2})?/gi,
            ""
        );

    // Remove repeated separators
    text =
        text.replace(
            /[|•]+/g,
            " "
        );

    // Remove excess punctuation
    text =
        text.replace(
            /\s*[-–—]\s*$/g,
            ""
        );

    text =
        text
            .replace(/\s+/g, " ")
            .trim();

    if (!text) {
        return "Bank transaction";
    }

    // Keep dashboard cards clean
    if (text.length > 55) {

        text =
            text.slice(0, 55).trim() +
            "…";

    }

    return text;
}

// ========================================
// CALCULATE FINANCIALS
// ========================================

function calculateFinancials(
    transactions
) {

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

    transactions.forEach(
        transaction => {

            const amount =
                Math.abs(
                    Number(
                        transaction.amount
                    ) || 0
                );

            if (
                transaction.type === "income"
            ) {

                income +=
                    amount;

            } else {

                expenses +=
                    amount;

                if (
                    categories[
                        transaction.category
                    ] !== undefined
                ) {

                    categories[
                        transaction.category
                    ] +=
                        amount;

                }

            }

        }
    );

    return {

        transactions,

        income,

        expenses,

        savings:
            income - expenses,

        balance:
            income - expenses,

        categories

    };
}

// ========================================
// EMPTY DATA
// ========================================

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
// MONEY PARSER
// ========================================

function parseMoney(value) {

    if (
        value === undefined ||
        value === null
    ) {

        return 0;
    }

    const raw =
        String(value)
            .trim();

    if (!raw) {
        return 0;
    }

    const negative =
        raw.includes("-") ||
        (
            raw.includes("(") &&
            raw.includes(")")
        );

    const number =
        parseFloat(
            raw.replace(
                /[₦$£NGN,\s()]/gi,
                ""
            )
        );

    if (
        isNaN(number)
    ) {

        return 0;
    }

    return negative
        ? -Math.abs(number)
        : number;
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
            "Make sure you're using a valid Kuda CSV, XLSX, XLS, or PDF statement.";

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
