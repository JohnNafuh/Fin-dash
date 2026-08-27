// ========================================
// FIN-DASH
// BANK STATEMENT ANALYZER + LIVE MARKET
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

        if (file) startAnalysis(file);

    });
}


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

    const fileName = document.getElementById("analyzingFileName");
    const progress = document.getElementById("analysisProgressBar");

    if (fileName) fileName.textContent = file.name;
    if (progress) progress.style.width = "0%";

    let value = 0;

    const loading = setInterval(() => {

        value += Math.random() * 10;

        if (value > 90) value = 90;

        if (progress) {
            progress.style.width = value + "%";
        }

    }, 250);


    analyzeStatement(file)
        .then(data => {

            clearInterval(loading);

            if (progress) progress.style.width = "100%";

            setTimeout(() => {

                updateDashboard(data);
                showComplete(data, file.name);

            }, 600);

        })
        .catch(error => {

            clearInterval(loading);

            console.error("Statement analysis error:", error);

            showUploadError();

        });
}


// ========================================
// READ STATEMENT
// ========================================

async function analyzeStatement(file) {

    const extension =
        file.name.split(".").pop().toLowerCase();

    if (extension === "csv") {
        return parseCSV(await file.text());
    }

    if (extension === "xlsx" || extension === "xls") {

        await loadXLSX();

        const buffer = await file.arrayBuffer();

        const workbook =
            XLSX.read(buffer, { type: "array" });

        const sheet =
            workbook.Sheets[workbook.SheetNames[0]];

        const rows =
            XLSX.utils.sheet_to_json(
                sheet,
                { header: 1, defval: "" }
            );

        return parseRows(rows);
    }

    if (extension === "pdf") {

        await loadPDFJS();

        const buffer = await file.arrayBuffer();

        const pdf =
            await pdfjsLib.getDocument({
                data: buffer
            }).promise;

        const pages = [];

        for (let i = 1; i <= pdf.numPages; i++) {

            const page = await pdf.getPage(i);
            const content = await page.getTextContent();

            pages.push(content.items);

        }

        return parseKudaPDF(pages);
    }

    throw new Error("Unsupported file type");
}


// ========================================
// CSV
// ========================================

function parseCSV(text) {

    const rows = text
        .split(/\r?\n/)
        .filter(row => row.trim())
        .map(row =>
            row.split(",").map(cell =>
                cell
                    .replace(/^"|"$/g, "")
                    .trim()
            )
        );

    return parseRows(rows);
}


// ========================================
// CSV / XLSX ROW PARSER
// ========================================

function parseRows(rows) {

    if (!rows.length) return emptyData();

    const headers =
        rows[0].map(h =>
            String(h).toLowerCase().trim()
        );

    const findColumn = (...names) =>
        headers.findIndex(h =>
            names.some(name => h.includes(name))
        );

    const descriptionIndex =
        findColumn(
            "description",
            "narration",
            "details",
            "merchant",
            "particular"
        );

    const amountIndex =
        findColumn("amount", "value");

    const typeIndex =
        findColumn("type", "transaction");

    const creditIndex =
        findColumn("credit", "deposit");

    const debitIndex =
        findColumn("debit", "withdraw");

    const dateIndex =
        findColumn("date", "transaction date");

    const transactions = [];


    for (let i = 1; i < rows.length; i++) {

        const row = rows[i];

        if (!row || !row.length) continue;

        const description =
            descriptionIndex >= 0
                ? String(row[descriptionIndex] || "")
                : "";

        let amount = 0;
        let type = "";


        if (
            creditIndex >= 0 &&
            String(row[creditIndex]).trim()
        ) {

            amount = Math.abs(
                parseMoney(row[creditIndex])
            );

            type = "income";

        } else if (
            debitIndex >= 0 &&
            String(row[debitIndex]).trim()
        ) {

            amount = Math.abs(
                parseMoney(row[debitIndex])
            );

            type = "expense";

        } else if (amountIndex >= 0) {

            const raw =
                parseMoney(row[amountIndex]);

            amount = Math.abs(raw);

            type =
                detectType(
                    typeIndex >= 0
                        ? String(row[typeIndex] || "")
                        : "",
                    description,
                    raw
                );
        }


        if (!amount || !description) continue;

        transactions.push({
            date:
                dateIndex >= 0
                    ? String(row[dateIndex] || "")
                    : "",
            description,
            amount,
            type,
            category:
                type === "expense"
                    ? categorize(description)
                    : "Income"
        });
    }


    return calculateFinancials(transactions);
}


// ========================================
// KUDA PDF PARSER
// ========================================

function parseKudaPDF(pages) {

    const transactions = [];

    for (const items of pages) {

        const rows = buildPDFRows(items);

        for (const row of rows) {

            const transaction =
                parseKudaRow(row);

            if (transaction) {
                transactions.push(transaction);
            }
        }
    }


    // Remove duplicate transactions
    const unique = [];

    const seen = new Set();

    transactions.forEach(transaction => {

        const key =
            [
                transaction.date,
                transaction.description,
                transaction.amount,
                transaction.type
            ].join("|");

        if (!seen.has(key)) {

            seen.add(key);
            unique.push(transaction);

        }

    });


    return calculateFinancials(unique);
}


// ========================================
// REBUILD PDF TEXT INTO REAL ROWS
// ========================================

function buildPDFRows(items) {

    const sorted =
        items
            .filter(item => item.str && item.str.trim())
            .map(item => ({
                text: item.str.trim(),
                x: item.transform[4],
                y: item.transform[5]
            }))
            .sort((a, b) => b.y - a.y || a.x - b.x);


    const rows = [];

    const tolerance = 3;


    sorted.forEach(item => {

        let row =
            rows.find(r =>
                Math.abs(r.y - item.y) <= tolerance
            );


        if (!row) {

            row = {
                y: item.y,
                items: []
            };

            rows.push(row);
        }

        row.items.push(item);

    });


    return rows
        .sort((a, b) => b.y - a.y)
        .map(row =>
            row.items
                .sort((a, b) => a.x - b.x)
                .map(item => item.text)
        );
}


// ========================================
// PARSE KUDA ROW
// ========================================

function parseKudaRow(row) {

    const text = row.join(" ").replace(/\s+/g, " ").trim();

    if (!text) return null;


    // Must contain a recognizable date.
    const dateMatch =
        text.match(
            /\b(\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4})\b/
        );


    if (!dateMatch) return null;


    // Ignore obvious statement headings / summaries.
    const lower = text.toLowerCase();

    const ignored =
        [
            "opening balance",
            "closing balance",
            "available balance",
            "statement period",
            "account number",
            "account name",
            "total credits",
            "total debits",
            "transaction history",
            "transaction date",
            "description",
            "debit",
            "credit",
            "balance"
        ];


    if (
        ignored.some(word =>
            lower.includes(word)
        )
    ) {
        return null;
    }


    const date = dateMatch[0];

    const afterDate =
        text
            .slice(
                dateMatch.index +
                date.length
            )
            .trim();


    /*
     * Find money values.
     * We take the final meaningful amount(s)
     * rather than treating every number as a transaction.
     */

    const amounts =
        afterDate.match(
            /(?:₦|NGN)?\s?-?\d[\d,]*(?:\.\d{2})?/g
        );


    if (!amounts || !amounts.length) {
        return null;
    }


    const values =
        amounts
            .map(parseMoney)
            .filter(v => Math.abs(v) > 0);


    if (!values.length) return null;


    /*
     * Kuda rows commonly contain description,
     * debit/credit and balance.
     *
     * The last number is usually balance,
     * so use the previous amount when possible.
     */

    let amount;

    if (values.length >= 2) {
        amount = Math.abs(values[values.length - 2]);
    } else {
        amount = Math.abs(values[0]);
    }


    if (!amount) return null;


    let description =
        afterDate
            .replace(
                /(?:₦|NGN)?\s?-?\d[\d,]*(?:\.\d{2})?/g,
                ""
            )
            .replace(/\s+/g, " ")
            .trim();


    if (!description) return null;


    const type =
        detectType(
            "",
            description,
            afterDate
        );


    return {

        date,

        description,

        amount,

        type,

        category:
            type === "expense"
                ? categorize(description)
                : "Income"

    };
}


// ========================================
// TRANSACTION TYPE
// ========================================

function detectType(type, description, amount) {

    const value =
        (
            type +
            " " +
            description
        ).toLowerCase();


    if (
        value.includes("credit") ||
        value.includes("deposit") ||
        value.includes("salary") ||
        value.includes("income") ||
        value.includes("refund") ||
        value.includes("received") ||
        value.includes("inflow") ||
        value.includes("money received") ||
        value.includes("cashback")
    ) {
        return "income";
    }


    if (
        value.includes("debit") ||
        value.includes("withdraw") ||
        value.includes("payment") ||
        value.includes("purchase") ||
        value.includes("pos") ||
        value.includes("transfer to") ||
        value.includes("sent") ||
        value.includes("outflow") ||
        value.includes("airtime") ||
        value.includes("data")
    ) {
        return "expense";
    }


    if (typeof amount === "number" && amount < 0) {
        return "expense";
    }


    return "expense";
}


// ========================================
// EXPENSE CATEGORIES
// ========================================

function categorize(description) {

    const text =
        String(description).toLowerCase();


    if (
        /food|restaurant|eatery|meal|pizza|chicken|burger|suya|groceries|market|supermarket|kuda food/i
            .test(text)
    ) {
        return "Food";
    }


    if (
        /uber|bolt|taxi|transport|bus|fuel|petrol|diesel|parking|ride/i
            .test(text)
    ) {
        return "Transport";
    }


    if (
        /electric|power|phcn|aedc|water|internet|wifi|subscription|netflix|dstv|gotv|bill|utility/i
            .test(text)
    ) {
        return "Bills";
    }


    if (
        /movie|cinema|spotify|music|game|gaming|concert|entertainment|club/i
            .test(text)
    ) {
        return "Entertainment";
    }


    if (
        /shop|shopping|jumia|amazon|clothing|fashion|store|purchase/i
            .test(text)
    ) {
        return "Shopping";
    }


    if (
        /airtime|data|mtn|airtel|glo|9mobile/i
            .test(text)
    ) {
        return "Bills";
    }


    if (
        /transfer|bank|fee|charge|commission|stamp duty/i
            .test(text)
    ) {
        return "Other";
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
        Other: 0
    };


    transactions.forEach(transaction => {

        if (transaction.type === "income") {

            income += transaction.amount;

        } else {

            expenses += transaction.amount;

            if (categories[transaction.category] !== undefined) {
                categories[transaction.category] +=
                    transaction.amount;
            }

        }

    });


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


    const raw = String(value).trim();

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


    if (isNaN(number)) return 0;


    return negative
        ? -Math.abs(number)
        : number;
}


// ========================================
// UPDATE DASHBOARD
// ========================================

function updateDashboard(data) {

    // Balance
    const balance =
        document.querySelector(".balance h2");

    if (balance) {
        balance.textContent =
            formatNaira(data.balance);
    }


    // Summary
    const cards =
        document.querySelectorAll(
            ".summary .card h3"
        );


    if (cards.length >= 3) {

        cards[0].textContent =
            formatNaira(data.income);

        cards[1].textContent =
            formatNaira(data.expenses);

        cards[2].textContent =
            formatNaira(data.savings);

    }


    updateTransactions(data.transactions);

    updateSpending(data);

    updateBudget(data);

    updateInsights(data);
}


// ========================================
// RECENT TRANSACTIONS
// ========================================

function updateTransactions(transactions) {

    const list =
        document.querySelector(
            ".transaction-list"
        );


    if (!list) return;


    if (!transactions.length) {

        list.innerHTML = `
            <div class="transaction empty-transaction">
                <p>No transactions found.</p>
            </div>
        `;

        return;
    }


    list.innerHTML =
        transactions
            .slice(-8)
            .reverse()
            .map(transaction => `

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
                                transaction.date || ""
                            )}
                        </small>

                    </div>

                    <strong class="${transaction.type}">

                        ${transaction.type === "income" ? "+" : "-"}

                        ${formatNaira(
                            transaction.amount
                        )}

                    </strong>

                </div>

            `)
            .join("");
}


function cleanDescription(description) {

    return String(description)
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 90);
}


// ========================================
// SPENDING OVERVIEW
// ========================================

function updateSpending(data) {

    const spending =
        document.querySelector(".spending");

    if (!spending) return;


    const items =
        spending.querySelectorAll(
            ".spending-item"
        );


    const categories = [
        "Food",
        "Transport",
        "Bills",
        "Entertainment"
    ];


    const max =
        Math.max(
            ...categories.map(
                category =>
                    data.categories[category] || 0
            ),
            1
        );


    items.forEach((item, index) => {

        const category =
            categories[index];

        if (!category) return;


        const amount =
            data.categories[category] || 0;


        const value =
            item.querySelector(
                "div:first-child span"
            );

        const bar =
            item.querySelector(
                ".progress-bar"
            );


        if (value) {
            value.textContent =
                formatNaira(amount);
        }


        if (bar) {

            bar.style.width =
                (
                    amount / max * 100
                ) + "%";

        }

    });
}


// ========================================
// MONTHLY BUDGET
// ========================================

function updateBudget(data) {

    const budget =
        document.querySelector(".budget");

    if (!budget) return;


    const amount =
        budget.querySelector(
            ".budget-amount h3"
        );

    const subtitle =
        budget.querySelector(
            ".budget-amount p"
        );

    const message =
        budget.querySelector(
            ".budget-message"
        );

    const progress =
        budget.querySelector(
            ".budget-progress-bar"
        );


    if (amount) {
        amount.textContent =
            formatNaira(data.expenses);
    }


    if (subtitle) {
        subtitle.textContent =
            `of ${formatNaira(data.expenses)} spent`;
    }


    if (progress) {
        progress.style.width =
            data.expenses > 0
                ? "100%"
                : "0%";
    }


    if (message) {

        if (data.expenses > 0) {

            message.textContent =
                `${data.transactions.length} transactions analyzed.`;

        } else {

            message.textContent =
                "No expenses found in this statement.";

        }

    }
}


// ========================================
// FINANCIAL INSIGHTS
// ========================================

function updateInsights(data) {

    const card =
        document.querySelector(".insight-card");

    if (!card) return;


    const title =
        card.querySelector("h3");

    const text =
        card.querySelector("p");


    if (!title || !text) return;


    const biggest =
        Object.entries(data.categories)
            .sort((a, b) => b[1] - a[1])[0];


    if (!data.transactions.length) {

        title.textContent =
            "No transactions found";

        text.textContent =
            "Try another statement or check that the statement contains transaction records.";

        return;
    }


    title.textContent =
        "Your finances have been analyzed";


    if (biggest && biggest[1] > 0) {

        text.textContent =
            `${biggest[0]} is currently your largest spending category at ${formatNaira(biggest[1])}. You have ${formatNaira(data.savings)} left after recorded income and expenses.`;

    } else {

        text.textContent =
            `Fin-dash analyzed ${data.transactions.length} transactions with ${formatNaira(data.income)} in income and ${formatNaira(data.expenses)} in expenses.`;

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
    }

    if (expenses) {
        expenses.textContent =
            formatNaira(data.expenses);
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
            "Make sure your statement is a supported CSV, XLSX, XLS, or PDF file.";
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
// FORMAT
// ========================================

function formatNaira(value) {

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
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}


// ========================================
// XLSX
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
// PDF.JS
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
                    new Error("PDF.js failed to load")
                );
                return;
            }


            window.pdfjsLib.GlobalWorkerOptions.workerSrc =
                "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

            resolve();

        };

        script.onerror = reject;

        document.head.appendChild(script);

    });
}


// ========================================
// LIVE MARKET
// ========================================

const marketBar =
    document.querySelector(".market-wrapper");

const marketToggle =
    document.getElementById("marketToggle");

const marketLoading =
    document.getElementById("marketLoading");

const marketData =
    document.getElementById("marketData");

const canvas =
    document.getElementById("marketChart");

let marketOpen = false;
let priceTimer;
let resizeTimer;


if (marketToggle) {

    marketToggle.addEventListener("click", () => {

        marketOpen = !marketOpen;


        if (marketOpen) {

            marketBar.style.height = "280px";
            marketToggle.textContent = "⌄";

            marketLoading.style.display = "flex";
            marketData.style.display = "none";


            setTimeout(() => {

                if (!marketOpen) return;

                marketLoading.style.display = "none";
                marketData.style.display = "block";

                drawFinancialChart();
                startMarketNumbers();

            }, 900);

        } else {

            marketBar.style.height = "55px";
            marketToggle.textContent = "⌃";

            marketLoading.style.display = "none";
            marketData.style.display = "none";

            stopMarket();

        }

    });

}


// ========================================
// MARKET CHART
// ========================================

let marketPoints = [
    48,46,47,44,45,42,43,40,42,39,
    41,37,38,35,36,33,35,32,34,31,
    33,29,31,28,30,27,29,25,27,24,
    26,22,25,21,23,19,22,18,20,17
];


function drawFinancialChart() {

    if (!canvas) return;


    const rect =
        canvas.getBoundingClientRect();

    const width =
        Math.max(rect.width, 1);

    const height =
        Math.max(rect.height, 1);

    const dpr =
        window.devicePixelRatio || 1;


    canvas.width = width * dpr;
    canvas.height = height * dpr;


    const ctx =
        canvas.getContext("2d");


    ctx.setTransform(
        dpr,0,0,dpr,0,0
    );

    ctx.clearRect(
        0,0,width,height
    );


    ctx.strokeStyle =
        "rgba(255,255,255,.055)";

    ctx.lineWidth = 1;


    for (let i = 1; i <= 4; i++) {

        const y =
            (height / 5) * i;

        ctx.beginPath();

        ctx.moveTo(0,y);
        ctx.lineTo(width,y);

        ctx.stroke();

    }


    const min =
        Math.min(...marketPoints) - 4;

    const max =
        Math.max(...marketPoints) + 4;

    const range =
        max - min;


    const points =
        marketPoints.map((value,index) => ({

            x:
                (index /
                (marketPoints.length - 1)) *
                width,

            y:
                height -
                ((value - min) / range) *
                (height - 14) -
                7

        }));


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
                (current.x + next.x) / 2;


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


    spline();

    ctx.lineTo(width,height);
    ctx.lineTo(0,height);
    ctx.closePath();


    const gradient =
        ctx.createLinearGradient(
            0,0,0,height
        );


    gradient.addColorStop(
        0,
        "rgba(108,60,255,.25)"
    );

    gradient.addColorStop(
        1,
        "rgba(108,60,255,0)"
    );


    ctx.fillStyle = gradient;
    ctx.fill();


    spline();

    ctx.strokeStyle =
        "rgba(145,112,255,.35)";

    ctx.lineWidth = 7;

    ctx.lineCap = "round";
    ctx.lineJoin = "round";

    ctx.shadowBlur = 16;
    ctx.shadowColor =
        "rgba(108,60,255,.8)";

    ctx.stroke();


    spline();

    ctx.shadowBlur = 0;

    ctx.strokeStyle =
        "#bda8ff";

    ctx.lineWidth = 2.4;

    ctx.stroke();


    const last =
        points[points.length - 1];


    ctx.beginPath();

    ctx.arc(
        last.x,
        last.y,
        3.5,
        0,
        Math.PI * 2
    );

    ctx.fillStyle = "#fff";

    ctx.shadowBlur = 12;
    ctx.shadowColor = "#bda8ff";

    ctx.fill();

}


// ========================================
// MARKET NUMBERS
// ========================================

function startMarketNumbers() {

    clearInterval(priceTimer);


    const price =
        document.getElementById("usdPrice");

    const change =
        document.getElementById("usdChange");


    if (!price || !change) return;


    let currentPrice = 1530;


    priceTimer =
        setInterval(() => {

            currentPrice +=
                (Math.random() - .46) * 2;


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
                    Math.random() * .7 + .1
                ).toFixed(2) +
                "%";


            const last =
                marketPoints[
                    marketPoints.length - 1
                ];


            marketPoints.push(
                last +
                (Math.random() - .47) * 5
            );


            if (marketPoints.length > 42) {
                marketPoints.shift();
            }


            drawFinancialChart();

        },3000);

}


function stopMarket() {

    clearInterval(priceTimer);

    priceTimer = null;
}


// ========================================
// RESIZE
// ========================================

window.addEventListener("resize", () => {

    clearTimeout(resizeTimer);

    resizeTimer =
        setTimeout(() => {

            if (marketOpen) {
                drawFinancialChart();
            }

        },150);

});
