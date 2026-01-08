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

    console.log("Requesting Graph API token...");
    Office.context.mailbox.getCallbackTokenAsync({ isRest: false }, async (result) => {
      console.log("Token request status:", result.status);

      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const accessToken = result.value;
        console.log("Graph access token retrieved successfully.");

        const encodedId = Office.context.mailbox.convertToRestId(
            item.itemId,
            Office.MailboxEnums.RestVersion.v2_0
        );
        const forwardUrl = `https://graph.microsoft.com/v1.0/me/messages/${encodedId}/forward`;


        console.log("Forward URL:", forwardUrl);

        const payload = {
          toRecipients: [{ emailAddress: { address: configuredEmail } }],
          comment: "Forwarded via Report button"
        };
        console.log("Payload:", payload);

        try {
          console.log("Sending forward request to Graph...");
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
        console.error("Failed to get Graph token:", result.error);
      }
    });
  } catch (err) {
    console.error("Unexpected error:", err);
  }
}
