// ========================================
// FIN-DASH LIVE MARKET
// ========================================

const marketBar = document.querySelector(".market-wrapper");
const marketToggle = document.getElementById("marketToggle");
const marketLoading = document.getElementById("marketLoading");
const marketData = document.getElementById("marketData");
const canvas = document.getElementById("marketChart");

let marketOpen = false;
let chartAnimation;
let priceTimer;
let resizeTimer;


// ========================================
// OPEN / CLOSE MARKET
// ========================================

if (marketToggle) {

    marketToggle.addEventListener("click", () => {

        marketOpen = !marketOpen;

        if (marketOpen) {

            marketBar.style.height = "280px";
            marketToggle.textContent = "⌄";

            marketLoading.style.display = "flex";
            marketData.style.display = "none";

            /*
             * Let the intersecting rings breathe
             * before revealing the market.
             */
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

            stopChart();
        }

    });

}


// ========================================
// FINANCIAL DATA
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


// ========================================
// DRAW FINANCIAL CHART
// ========================================

function drawFinancialChart() {

    if (!canvas) return;

    stopChart();

    const rect = canvas.getBoundingClientRect();

    const width = Math.max(rect.width, 1);
    const height = Math.max(rect.height, 1);

    const dpr = window.devicePixelRatio || 1;

    canvas.width = width * dpr;
    canvas.height = height * dpr;

    const ctx = canvas.getContext("2d");

    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    /*
     * Clear
     */

    ctx.clearRect(0, 0, width, height);


    // ====================================
    // GRID
    // ====================================

    ctx.save();

    ctx.strokeStyle = "rgba(255,255,255,0.055)";
    ctx.lineWidth = 1;

    const horizontalLines = 4;

    for (let i = 1; i <= horizontalLines; i++) {

        const y =
            (height / (horizontalLines + 1)) * i;

        ctx.beginPath();

        ctx.moveTo(0, y);
        ctx.lineTo(width, y);

        ctx.stroke();

    }

    ctx.restore();


    // ====================================
    // POINTS
    // ====================================

    const min = Math.min(...marketPoints) - 4;
    const max = Math.max(...marketPoints) + 4;

    const range = max - min;

    const points = marketPoints.map((value, index) => {

        const x =
            (index / (marketPoints.length - 1)) *
            width;

        const y =
            height -
            ((value - min) / range) *
            (height - 14) -
            7;

        return { x, y };

    });


    // ====================================
    // SMOOTH SPLINE
    // ====================================

    function drawSplinePath() {

        ctx.beginPath();

        ctx.moveTo(
            points[0].x,
            points[0].y
        );

        for (let i = 0; i < points.length - 1; i++) {

            const current = points[i];
            const next = points[i + 1];

            const midpointX =
                (current.x + next.x) / 2;

            ctx.bezierCurveTo(
                midpointX,
                current.y,
                midpointX,
                next.y,
                next.x,
                next.y
            );

        }

    }


    // ====================================
    // AREA FILL
    // ====================================

    ctx.save();

    drawSplinePath();

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
        "rgba(108,60,255,0.25)"
    );

    gradient.addColorStop(
        1,
        "rgba(108,60,255,0)"
    );

    ctx.fillStyle = gradient;

    ctx.fill();

    ctx.restore();


    // ====================================
    // GLOW
    // ====================================

    ctx.save();

    drawSplinePath();

    ctx.strokeStyle =
        "rgba(145,112,255,0.35)";

    ctx.lineWidth = 7;

    ctx.lineCap = "round";
    ctx.lineJoin = "round";

    ctx.shadowBlur = 16;
    ctx.shadowColor =
        "rgba(108,60,255,0.8)";

    ctx.stroke();

    ctx.restore();


    // ====================================
    // MAIN LINE
    // ====================================

    ctx.save();

    drawSplinePath();

    ctx.strokeStyle = "#bda8ff";

    ctx.lineWidth = 2.4;

    ctx.lineCap = "round";
    ctx.lineJoin = "round";

    ctx.stroke();

    ctx.restore();


    // ====================================
    // LAST PRICE POINT
    // ====================================

    const last =
        points[points.length - 1];

    ctx.save();

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

    ctx.restore();

}


// ========================================
// SIMULATED MARKET MOVEMENT
// ========================================

function startMarketNumbers() {

    if (priceTimer) {
        clearInterval(priceTimer);
    }

    const price =
        document.getElementById("usdPrice");

    const change =
        document.getElementById("usdChange");

    if (!price || !change) return;


    let currentPrice = 1530;


    priceTimer = setInterval(() => {

        /*
         * Small movements so the market
         * doesn't jump unrealistically.
         */

        const movement =
            (Math.random() - 0.46) * 2;

        currentPrice += movement;


        price.textContent =
            "₦" +
            currentPrice.toLocaleString(
                "en-NG",
                {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2
                }
            );


        const percentage =
            (
                Math.random() * 0.7 +
                0.1
            ).toFixed(2);


        change.textContent =
            "+" + percentage + "%";


        /*
         * Push a new point into the chart.
         */

        const last =
            marketPoints[marketPoints.length - 1];

        const next =
            last +
            (Math.random() - 0.47) * 5;


        marketPoints.push(next);

        /*
         * Keep the chart from becoming
         * infinitely long.
         */

        if (marketPoints.length > 42) {
            marketPoints.shift();
        }


        drawFinancialChart();

    }, 3000);

}


// ========================================
// STOP MARKET
// ========================================

function stopChart() {

    if (chartAnimation) {

        cancelAnimationFrame(
            chartAnimation
        );

        chartAnimation = null;

    }

    if (priceTimer) {

        clearInterval(priceTimer);

        priceTimer = null;

    }

}


// ========================================
// RESIZE
// ========================================

window.addEventListener(
    "resize",
    () => {

        clearTimeout(resizeTimer);

        resizeTimer = setTimeout(() => {

            if (marketOpen) {
                drawFinancialChart();
            }

        }, 150);

    }
);
