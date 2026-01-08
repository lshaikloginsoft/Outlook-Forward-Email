Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Office.js is ready. Host:", info.host);
    document.getElementById("reportButton").onclick = forwardEmail;
  }
});

async function forwardEmail() {
  const configuredEmail = "lshaik@loginsoft.com"; // change to your saved variable
  console.log("Forward button clicked. Target email:", configuredEmail);

  try {
    const item = Office.context.mailbox.item;
    console.log("Selected item ID:", item.itemId);

    if (!confirm("Forward this email to " + configuredEmail + "?")) {
      console.log("User cancelled forwarding.");
      return;
    }

    console.log("Requesting callback token for REST API...");
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
      console.log("Token request status:", result.status);

      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const accessToken = result.value;
        console.log("Access token retrieved successfully.");

        const restUrl = Office.context.mailbox.restUrl;
        const itemId = item.itemId;
        const forwardUrl = `${restUrl}/v2.0/me/messages/${itemId}/forward`;

        console.log("Forward URL:", forwardUrl);

        const payload = {
          ToRecipients: [{ EmailAddress: { Address: configuredEmail } }],
          Comment: "Forwarded via Report button"
        };
        console.log("Payload:", payload);

        try {
          console.log("Sending forward request...");
          const response = await fetch(forwardUrl, {
            method: "POST",
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: "application/json",
              "Content-Type": "application/json"
            },
            body: JSON.stringify(payload)
          });

          console.log("Response status:", response.status);

          if (response.ok) {
            console.log("Forward request succeeded.");
            alert("Email forwarded successfully.");
          } else {
            const error = await response.json();
            console.error("Forward request failed:", error);
            alert("Failed: " + error.error.message);
          }
        } catch (fetchErr) {
          console.error("Error during fetch:", fetchErr);
          alert("Network error occurred.");
        }
      } else {
        console.error("Failed to get access token:", result.error);
        alert("Could not get access token.");
      }
    });
  } catch (err) {
    console.error("Unexpected error:", err);
    alert("Unexpected error occurred.");
  }
}
