// ========================================
// FIN-DASH
// STATEMENT ANALYZER + LIVE MARKET
// ========================================


// ========================================
// STATEMENT ANALYZER
// ========================================

const analyzer = document.querySelector(".statement-analyzer");
const uploadButton = document.querySelector(".upload-button");

let statementInput;


// Create file picker automatically
if (uploadButton) {

    statementInput = document.createElement("input");

    statementInput.type = "file";
    statementInput.accept = ".csv,.xlsx,.pdf";
    statementInput.style.display = "none";

    document.body.appendChild(statementInput);


    uploadButton.addEventListener("click", () => {
        statementInput.click();
    });


    statementInput.addEventListener("change", async () => {

        const file = statementInput.files[0];

        if (!file) return;

        startStatementAnalysis(file);

    });

}


// ========================================
// START ANALYSIS
// ========================================

function startStatementAnalysis(file) {

    if (!analyzer) return;

    analyzer.innerHTML = `
        <div class="statement-analyzing">

            <div class="statement-rings">
                <span></span>
                <span></span>
                <span></span>
            </div>

            <h2>Statement uploaded</h2>

            <p id="analysisText">
                Analyzing your finances...
            </p>

            <div class="analysis-progress">
                <div
                    class="analysis-progress-bar"
                    id="analysisProgress"
                ></div>
            </div>

        </div>
    `;


    const progress =
        document.getElementById("analysisProgress");

    let value = 0;


    const loading = setInterval(() => {

        value += Math.random() * 12;

        if (value > 92) {
            value = 92;
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
                showAnalysisComplete(data);

            }, 500);

        })
        .catch(error => {

            clearInterval(loading);

            console.error(error);

            analyzer.innerHTML = `
                <div class="statement-upload">

                    <div class="upload-icon">!</div>

                    <h2>Couldn't analyze statement</h2>

                    <p>
                        Make sure your statement is a
                        supported CSV, XLSX, or PDF file.
                    </p>

                    <button
                        class="upload-button"
                        onclick="location.reload()"
                    >
                        Try Again
                    </button>

                </div>
            `;

        });

}


// ========================================
// READ STATEMENT
// ========================================

async function analyzeStatement(file) {

    const extension =
        file.name.split(".").pop().toLowerCase();


    // CSV
    if (extension === "csv") {

        const text = await file.text();

        return parseCSV(text);

    }


    // XLSX
    if (extension === "xlsx") {

        await loadXLSX();

        const buffer = await file.arrayBuffer();

        const workbook =
            XLSX.read(buffer, { type: "array" });

        const sheet =
            workbook.Sheets[workbook.SheetNames[0]];

        const rows =
            XLSX.utils.sheet_to_json(
                sheet,
                { header: 1 }
            );

        return parseRows(rows);

    }


    // PDF
    if (extension === "pdf") {

        await loadPDFJS();

        const buffer = await file.arrayBuffer();

        const pdf =
            await pdfjsLib.getDocument({
                data: buffer
            }).promise;

        let text = "";

        for (let i = 1; i <= pdf.numPages; i++) {

            const page =
                await pdf.getPage(i);

            const content =
                await page.getTextContent();

            text +=
                content.items
                    .map(item => item.str)
                    .join(" ") + "\n";
        }

        return parsePDFText(text);

    }


    throw new Error("Unsupported file");

}


// ========================================
// CSV
// ========================================

function parseCSV(text) {

    const rows =
        text
            .split(/\r?\n/)
            .filter(row => row.trim())
            .map(row =>
                row.split(",").map(cell =>
                    cell.replace(/^"|"$/g, "").trim()
                )
            );

    return parseRows(rows);

}


// ========================================
// GENERAL ROW PARSER
// ========================================

function parseRows(rows) {

    if (!rows.length) {
        return emptyData();
    }


    const headers =
        rows[0].map(h =>
            String(h).toLowerCase()
        );


    const transactions = [];


    for (let i = 1; i < rows.length; i++) {

        const row = rows[i];

        if (!row || !row.length) continue;


        const descriptionIndex =
            headers.findIndex(h =>
                h.includes("description") ||
                h.includes("narration") ||
                h.includes("details") ||
                h.includes("merchant")
            );


        const amountIndex =
            headers.findIndex(h =>
                h.includes("amount") ||
                h.includes("value")
            );


        const typeIndex =
            headers.findIndex(h =>
                h.includes("type") ||
                h.includes("transaction")
            );


        const creditIndex =
            headers.findIndex(h =>
                h.includes("credit") ||
                h.includes("deposit")
            );


        const debitIndex =
            headers.findIndex(h =>
                h.includes("debit") ||
                h.includes("withdraw")
            );


        const description =
            descriptionIndex >= 0
                ? String(row[descriptionIndex] || "")
                : row.join(" ");


        let amount = 0;
        let type = "";


        if (creditIndex >= 0 && row[creditIndex]) {

            amount =
                parseMoney(row[creditIndex]);

            type = "credit";

        } else if (debitIndex >= 0 && row[debitIndex]) {

            amount =
                parseMoney(row[debitIndex]);

            type = "debit";

        } else if (amountIndex >= 0) {

            amount =
                parseMoney(row[amountIndex]);

            type =
                typeIndex >= 0
                    ? String(row[typeIndex]).toLowerCase()
                    : "";

        }


        if (!amount) continue;


        transactions.push({
            description,
            amount: Math.abs(amount),
            type: detectType(type, description, amount)
        });

    }


    return calculateFinancials(transactions);

}


// ========================================
// PDF PARSER
// ========================================

function parsePDFText(text) {

    const transactions = [];

    const lines =
        text
            .split(/\r?\n/)
            .map(line => line.trim())
            .filter(Boolean);


    for (const line of lines) {

        const numbers =
            line.match(/(?:₦|\$)?\s?[\d,]+(?:\.\d{2})?/g);

        if (!numbers || !numbers.length) continue;


        const amount =
            parseMoney(numbers[numbers.length - 1]);


        if (!amount) continue;


        transactions.push({
            description: line,
            amount: Math.abs(amount),
            type: detectType("", line, amount)
        });

    }


    return calculateFinancials(transactions);

}


// ========================================
// DETECT INCOME / EXPENSE
// ========================================

function detectType(type, description, amount) {

    const value =
        (type + " " + description).toLowerCase();


    if (
        value.includes("credit") ||
        value.includes("deposit") ||
        value.includes("salary") ||
        value.includes("income") ||
        value.includes("transfer from") ||
        value.includes("refund")
    ) {
        return "income";
    }


    if (
        value.includes("debit") ||
        value.includes("withdraw") ||
        value.includes("payment") ||
        value.includes("purchase") ||
        value.includes("pos") ||
        value.includes("transfer to")
    ) {
        return "expense";
    }


    return amount < 0
        ? "expense"
        : "expense";

}


// ========================================
// CALCULATE FINANCIALS
// ========================================

function calculateFinancials(transactions) {

    let income = 0;
    let expenses = 0;


    transactions.forEach(transaction => {

        if (transaction.type === "income") {
            income += transaction.amount;
        } else {
            expenses += transaction.amount;
        }

    });


    return {
        transactions,
        income,
        expenses,
        savings: income - expenses,
        balance: income - expenses
    };

}


function emptyData() {

    return {
        transactions: [],
        income: 0,
        expenses: 0,
        savings: 0,
        balance: 0
    };

}


// ========================================
// MONEY PARSER
// ========================================

function parseMoney(value) {

    if (value === undefined || value === null) {
        return 0;
    }


    let number =
        String(value)
            .replace(/[₦$£,\s]/g, "")
            .replace(/[()]/g, "");


    return parseFloat(number) || 0;

}


// ========================================
// UPDATE DASHBOARD
// ========================================

function updateDashboard(data) {

    const balance =
        document.querySelector(".balance h2");

    const cards =
        document.querySelectorAll(".summary .card h3");


    if (balance) {
        balance.textContent =
            formatNaira(data.balance);
    }


    if (cards.length >= 3) {

        cards[0].textContent =
            formatNaira(data.income);

        cards[1].textContent =
            formatNaira(data.expenses);

        cards[2].textContent =
            formatNaira(data.savings);

    }


    // Recent transactions
    const list =
        document.querySelector(".transaction-list");


    if (list && data.transactions.length) {

        list.innerHTML =
            data.transactions
                .slice(-8)
                .reverse()
                .map(transaction => `

                    <div class="transaction">

                        <div>
                            <strong>
                                ${escapeHTML(
                                    transaction.description
                                )}
                            </strong>
                        </div>

                        <strong>
                            ${transaction.type === "income" ? "+" : "-"}
                            ${formatNaira(transaction.amount)}
                        </strong>

                    </div>

                `)
                .join("");

    }

}


// ========================================
// ANALYSIS COMPLETE
// ========================================

function showAnalysisComplete(data) {

    analyzer.innerHTML = `

        <div class="statement-complete">

            <div class="complete-icon">
                ✓
            </div>

            <h2>Analysis complete</h2>

            <p>
                Your financial dashboard has been updated.
            </p>

            <div class="statement-stats">

                <div>
                    <strong>
                        ${data.transactions.length}
                    </strong>
                    <span>Transactions</span>
                </div>

                <div>
                    <strong>
                        ${formatNaira(data.income)}
                    </strong>
                    <span>Income</span>
                </div>

                <div>
                    <strong>
                        ${formatNaira(data.expenses)}
                    </strong>
                    <span>Expenses</span>
                </div>

            </div>

            <button
                class="analyze-another"
                onclick="location.reload()"
            >
                Analyze Another Statement
            </button>

        </div>

    `;

}


// ========================================
// FORMAT CURRENCY
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
            "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs";

        script.type = "module";


        script.onload = async () => {

            try {

                window.pdfjsLib =
                    await import(script.src);

                resolve();

            } catch (error) {

                reject(error);

            }

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
    48, 46, 47, 44, 45,
    42, 43, 40, 42, 39,
    41, 37, 38, 35, 36,
    33, 35, 32, 34, 31,
    33, 29, 31, 28, 30,
    27, 29, 25, 27, 24,
    26, 22, 25, 21, 23,
    19, 22, 18, 20, 17
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


    canvas.width =
        width * dpr;

    canvas.height =
        height * dpr;


    const ctx =
        canvas.getContext("2d");

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


    // Grid
    ctx.strokeStyle =
        "rgba(255,255,255,0.055)";

    ctx.lineWidth = 1;


    for (let i = 1; i <= 4; i++) {

        const y =
            (height / 5) * i;

        ctx.beginPath();

        ctx.moveTo(0, y);
        ctx.lineTo(width, y);

        ctx.stroke();

    }


    const min =
        Math.min(...marketPoints) - 4;

    const max =
        Math.max(...marketPoints) + 4;

    const range =
        max - min;


    const points =
        marketPoints.map((value, index) => ({

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


    // Area
    spline();

    ctx.lineTo(width, height);
    ctx.lineTo(0, height);
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

    ctx.fillStyle = gradient;
    ctx.fill();


    // Glow
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


    // Main line
    spline();

    ctx.shadowBlur = 0;

    ctx.strokeStyle =
        "#bda8ff";

    ctx.lineWidth = 2.4;

    ctx.stroke();


    // Last point
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

    ctx.fillStyle = "#ffffff";

    ctx.shadowBlur = 12;
    ctx.shadowColor = "#bda8ff";

    ctx.fill();

}


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
                (Math.random() - .47) * 5
            );


            if (marketPoints.length > 42) {
                marketPoints.shift();
            }


            drawFinancialChart();

        }, 3000);

}


function stopMarket() {

    clearInterval(priceTimer);

    priceTimer = null;

}


window.addEventListener(
    "resize",
    () => {

        clearTimeout(resizeTimer);

        resizeTimer =
            setTimeout(() => {

                if (marketOpen) {
                    drawFinancialChart();
                }

            }, 150);

    }
);
