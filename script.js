// ========================================
// FIN-DASH
// STATEMENT ANALYZER + LIVE MARKET
// ========================================


// ========================================
// STATEMENT ANALYZER
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

        if (!file) return;

        startAnalysis(file);
    });
}


// Analyze another statement
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

        value += Math.random() * 10;

        if (value > 90) value = 90;

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

        const text = await file.text();

        return parseCSV(text);
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

        let text = "";

        for (let i = 1; i <= pdf.numPages; i++) {

            const page = await pdf.getPage(i);

            const content =
                await page.getTextContent();

            text +=
                content.items
                    .map(item => item.str)
                    .join(" ") + "\n";
        }

        return parsePDFText(text);
    }

    throw new Error("Unsupported file type");
}


// ========================================
// CSV PARSER
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
// ROW PARSER
// ========================================

function parseRows(rows) {

    if (!rows.length) {
        return emptyData();
    }

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
        findColumn(
            "amount",
            "value"
        );

    const typeIndex =
        findColumn(
            "type",
            "transaction type"
        );

    const creditIndex =
        findColumn(
            "credit",
            "deposit"
        );

    const debitIndex =
        findColumn(
            "debit",
            "withdraw"
        );

    const transactions = [];


    for (let i = 1; i < rows.length; i++) {

        const row = rows[i];

        if (!row || !row.length) continue;

        const description =
            descriptionIndex >= 0
                ? String(row[descriptionIndex] || "")
                : row.join(" ");


        let amount = 0;
        let type = "";


        if (
            creditIndex >= 0 &&
            String(row[creditIndex]).trim()
        ) {

            amount =
                parseMoney(row[creditIndex]);

            type = "credit";

        } else if (
            debitIndex >= 0 &&
            String(row[debitIndex]).trim()
        ) {

            amount =
                parseMoney(row[debitIndex]);

            type = "debit";

        } else if (amountIndex >= 0) {

            amount =
                parseMoney(row[amountIndex]);

            type =
                typeIndex >= 0
                    ? String(row[typeIndex] || "").toLowerCase()
                    : "";
        }


        if (!amount) continue;


        transactions.push({

            description,

            amount: Math.abs(amount),

            type:
                detectType(
                    type,
                    description,
                    amount
                )

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
            line.match(
                /(?:₦|\$|£)?\s?-?[\d,]+(?:\.\d{2})?/g
            );

        if (!numbers) continue;


        const rawAmount =
            numbers[numbers.length - 1];

        const amount =
            parseMoney(rawAmount);

        if (!amount) continue;


        transactions.push({

            description: line,

            amount: Math.abs(amount),

            type:
                detectType(
                    "",
                    line,
                    amount
                )

        });
    }


    return calculateFinancials(transactions);
}


// ========================================
// DETECT TRANSACTION TYPE
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
        value.includes("inflow")
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
        value.includes("outflow")
    ) {
        return "expense";
    }


    return "expense";
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

        savings:
            income - expenses,

        balance:
            income - expenses

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
// MONEY
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
                /[₦$£,\s()]/g,
                ""
            )
        );


    if (isNaN(number)) {
        return 0;
    }


    return negative
        ? -Math.abs(number)
        : number;
}


// ========================================
// UPDATE DASHBOARD
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
    }


    if (cards.length >= 3) {

        cards[0].textContent =
            formatNaira(data.income);

        cards[1].textContent =
            formatNaira(data.expenses);

        cards[2].textContent =
            formatNaira(data.savings);
    }


    const list =
        document.querySelector(
            ".transaction-list"
        );


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
                            ${formatNaira(
                                transaction.amount
                            )}
                        </strong>

                    </div>

                `)
                .join("");
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


    if (name) {
        name.textContent = fileName;
    }

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
// FORMAT NAIRA
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
            "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";

        script.onload = () => {

            if (window.pdfjsLib) {

                window.pdfjsLib.GlobalWorkerOptions.workerSrc =
                    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

                resolve();

            } else {

                reject(
                    new Error("PDF.js failed to load")
                );

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
        dpr, 0, 0, dpr, 0, 0
    );

    ctx.clearRect(
        0, 0, width, height
    );


    // Grid

    ctx.strokeStyle =
        "rgba(255,255,255,.055)";

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
            0, 0, 0, height
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


    priceTimer = setInterval(() => {

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

    }, 3000);
}


// ========================================
// STOP MARKET
// ========================================

function stopMarket() {

    clearInterval(priceTimer);

    priceTimer = null;
}


// ========================================
// RESIZE
// ========================================

window.addEventListener("resize", () => {

    clearTimeout(resizeTimer);

    resizeTimer = setTimeout(() => {

        if (marketOpen) {
            drawFinancialChart();
        }

    }, 150);

});
