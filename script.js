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
let chartProgress = 0;


// ========================================
// OPEN / CLOSE MARKET BAR
// ========================================

marketToggle.addEventListener("click", () => {

    marketOpen = !marketOpen;

    if (marketOpen) {

        marketBar.style.height = "280px";
        marketToggle.textContent = "⌄";

        marketLoading.style.display = "flex";
        marketData.style.display = "none";

        // Start loading animation
        setTimeout(() => {

            marketLoading.style.display = "none";
            marketData.style.display = "block";

            startChart();

        }, 2200);

    } else {

        marketBar.style.height = "55px";
        marketToggle.textContent = "⌃";

        marketLoading.style.display = "none";
        marketData.style.display = "none";

        stopChart();
    }

});


// ========================================
// MARKET CHART
// ========================================

function startChart() {

    if (!canvas) return;

    const ctx = canvas.getContext("2d");

    const width = canvas.width = canvas.offsetWidth * window.devicePixelRatio;
    const height = canvas.height = canvas.offsetHeight * window.devicePixelRatio;

    ctx.scale(window.devicePixelRatio, window.devicePixelRatio);

    const chartWidth = canvas.offsetWidth;
    const chartHeight = canvas.offsetHeight;

    const points = [];

    for (let i = 0; i < 40; i++) {

        const wave =
            Math.sin(i * 0.55) * 10 +
            Math.sin(i * 0.19) * 7;

        const noise = Math.random() * 8;

        points.push(
            chartHeight / 2 - wave - noise
        );
    }

    chartProgress = 0;

    function drawChart() {

        ctx.clearRect(0, 0, chartWidth, chartHeight);

        ctx.beginPath();

        const visiblePoints =
            Math.floor(points.length * chartProgress);

        for (let i = 0; i < visiblePoints; i++) {

            const x =
                (i / (points.length - 1)) *
                chartWidth;

            const y = points[i];

            if (i === 0) {
                ctx.moveTo(x, y);
            } else {
                ctx.lineTo(x, y);
            }

        }

        ctx.strokeStyle = "#bda8ff";
        ctx.lineWidth = 2;
        ctx.lineJoin = "round";
        ctx.lineCap = "round";

        ctx.stroke();


        // Glow

        ctx.shadowBlur = 10;
        ctx.shadowColor = "#8d6cff";

        ctx.stroke();

        ctx.shadowBlur = 0;


        // Continue animation

        chartProgress += 0.025;

        if (chartProgress < 1) {

            chartAnimation =
                requestAnimationFrame(drawChart);

        } else {

            animateMarketData();

        }

    }

    drawChart();
}


// ========================================
// LIVE NUMBER ANIMATION
// ========================================

function animateMarketData() {

    const price =
        document.getElementById("usdPrice");

    const change =
        document.getElementById("usdChange");

    if (!price || !change) return;

    let current = 1530;

    setInterval(() => {

        const movement =
            (Math.random() - 0.45) * 2;

        current += movement;

        price.textContent =
            "₦" +
            current.toLocaleString("en-NG", {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });

        const percentage =
            (Math.random() * 0.8 + 0.1).toFixed(2);

        change.textContent =
            "+" + percentage + "%";

    }, 2500);

}


// ========================================
// STOP CHART
// ========================================

function stopChart() {

    if (chartAnimation) {

        cancelAnimationFrame(chartAnimation);

    }

}


// ========================================
// RESPONSIVE CANVAS
// ========================================

window.addEventListener("resize", () => {

    if (marketOpen) {

        startChart();

    }

});
