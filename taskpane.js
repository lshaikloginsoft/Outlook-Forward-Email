Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Office.js is ready. Host:", info.host);
    document.getElementById("reportButton").onclick = forwardEmail;
  }
});

async function forwardEmail() {
  const configuredEmail = "lshaik@loginsoft.com"; // your target email
  console.log("Forward button clicked. Target email:", configuredEmail);

  try {
    const item = Office.context.mailbox.item;
    console.log("Selected item ID:", item.itemId);

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
          } else {
            const error = await response.json();
            console.error("Forward request failed:", error);
          }
        } catch (fetchErr) {
          console.error("Error during fetch:", fetchErr);
        }
      } else {
        console.error("Failed to get access token:", result.error);
      }
    });
  } catch (err) {
    console.error("Unexpected error:", err);
  }
}
