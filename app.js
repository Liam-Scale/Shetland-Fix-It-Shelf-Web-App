async function login() {
    const tenantId = "9bcad0a7-b9ab-4a75-b321-5c55972f0032";       // paste your Tenant ID here
    const clientId = "a6a622dc-e69a-4db7-aa97-4ffe9f3cf2ef";       // paste your App Registration client ID here
    const redirectUri = encodeURIComponent("https://liam-scale.github.io/Shetland-Fix-It-Shelf-Web-App/");
    const scope = "Files.ReadWrite.All";

    const loginUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}&scope=${scope}`;

    window.location.href = loginUrl;
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


