const SCORE_FIELDS = [
    "Comfort",
    "Security",
    "Case & Battery",
    "Controls",
    "Connectivity",
    "Call Quality",
    "App & Extras",
    "Noise Cancelling",
    "Transparency",
    "Sound Quality"
];

let earbuds = [];

// Fetch the .xlsx file directly from the repo directory
fetch("earbuds_scores v2.xlsx")
    .then(response => {
        if (!response.ok) {
            throw new Error("Failed to load Excel file.");
        }
        return response.arrayBuffer();
    })
    .then(buffer => {
        // Read the binary spreadsheet
        const workbook = XLSX.read(buffer, { type: "array" });
        
        // Grab the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to JSON using default settings so Row 1 sets the correct headers 
        // (e.g. "Earbuds", "price.com", "Comfort")
        const rawData = XLSX.utils.sheet_to_json(worksheet);
        
        earbuds = rawData;
        renderList(earbuds);
    })
    .catch(error => {
        console.error("Error loading or parsing the Excel file:", error);
        document.getElementById("list").textContent = "Error loading spreadsheet data. Make sure 'earbuds_scores v2.xlsx' is uploaded.";
    });

function renderList(items) {
    const list = document.getElementById("list");
    list.innerHTML = "";

    // Filter out rows that don't have an earbud name, 
    // AND skip the custom row where the name is literally "weight"
    const validItems = items.filter(item => {
        return item.Earbuds && item.Earbuds.toString().trim().toLowerCase() !== "weight";
    });

    if (validItems.length === 0) {
        list.innerHTML = "No entries found in spreadsheet.";
        return;
    }

    validItems.forEach(item => {
        const displayPrice = item["price.com"] ? "HK$ " + item["price.com"] : "N/A";
        const displayScore = item["Total Score"] !== undefined ? item["Total Score"] : "--";

        const card = document.createElement("div");
        card.className = "card";
        card.innerHTML = `
            <div class="card-left">
                <h3>${item.Earbuds}</h3>
                <div class="price">${displayPrice}</div>
            </div>
            <div class="card-right">
                <div class="score">${displayScore}</div>
                <div class="label">score</div>
            </div>
            <div class="arrow">›</div>
        `;
        card.onclick = () => showDetail(item);
        list.appendChild(card);
    });
}

function showDetail(item) {
    document.getElementById("detail-title").textContent = item.Earbuds;
    
    const displayPrice = item["price.com"] ? "Price: HK$ " + item["price.com"] : "Price: N/A";
    document.getElementById("detail-price").textContent = displayPrice;

    const container = document.getElementById("detail-scores");
    container.innerHTML = "";

    SCORE_FIELDS.forEach(field => {
        if (item[field] !== undefined) {
            const row = document.createElement("div");
            row.className = "score-row";
            row.innerHTML = `
                <span class="score-label">${field}</span>
                <span class="score-val">${item[field]}</span>
            `;
            container.appendChild(row);
        }
    });

    // Total Score Row
    if (item["Total Score"] !== undefined) {
        const total = document.createElement("div");
        total.className = "total-row";
        total.innerHTML = `
            <span class="total-label">Total Score</span>
            <span class="total-val">${item["Total Score"]}</span>
        `;
        container.appendChild(total);
    }

    // Weighted Score Row
    if (item["weighted score"] !== undefined) {
        const weighted = document.createElement("div");
        weighted.className = "score-row";
        weighted.innerHTML = `
            <span class="score-label">Weighted Score</span>
            <span class="score-val">${item["weighted score"]}</span>
        `;
        container.appendChild(weighted);
    }

    document.getElementById("detail").style.display = "block";
}

function hideDetail() {
    document.getElementById("detail").style.display = "none";
}