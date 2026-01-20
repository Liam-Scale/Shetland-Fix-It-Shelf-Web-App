async function login() {
    // Redirect to Microsoft login for SPA
    window.location.href = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=a6a622dc-e69a-4db7-aa97-4ffe9f3cf2ef&response_type=token&redirect_uri=" + encodeURIComponent(window.location.href) + "&scope=Files.ReadWrite.All";
}

window.onload = () => {
    // After login, parse token from URL (if present)
    const hash = window.location.hash;
    if (hash.includes("access_token")) {
        document.querySelector("button").hidden = true;
        document.getElementById("repairForm").hidden = false;
        const token = new URLSearchParams(hash.substring(1)).get("access_token");
        window.accessToken = token; // save for Graph calls
    }
};

document.getElementById("repairForm").addEventListener("submit", async (e) => {
    e.preventDefault();

    const payload = {
        customer: customer.value,
        equipment: equipment.value,
        serial: serial.value,
        fault: fault.value
    };

    const excelUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/Fix It Shelf Log.xlsx:/workbook/tables/FixItShelf/rows/add";

    const body = {
        values: [[
            new Date().toISOString(),
            payload.customer,
            payload.equipment,
            payload.serial,
            payload.fault,
            "Received"
        ]]
    };

    const res = await fetch(excelUrl, {
        method: "POST",
        headers: {
            Authorization: `Bearer ${window.accessToken}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });

    status.textContent = res.ok ? "Saved to Excel ✔️" : "Error saving";
});