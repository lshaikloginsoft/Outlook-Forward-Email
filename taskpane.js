Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("reportButton").onclick = forwardEmail;
  }
});

async function forwardEmail() {
  const configuredEmail = "lshaik@loginsoft.com"; // change to your saved variable

  try {
    const item = Office.context.mailbox.item;

    if (!confirm("Forward this email to " + configuredEmail + "?")) return;

    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const accessToken = result.value;
        const restUrl = Office.context.mailbox.restUrl;
        const itemId = item.itemId;

        const forwardUrl = `${restUrl}/v2.0/me/messages/${itemId}/forward`;

        const payload = {
          ToRecipients: [{ EmailAddress: { Address: configuredEmail } }],
          Comment: "Forwarded via Report button"
        };

        const response = await fetch(forwardUrl, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: "application/json",
            "Content-Type": "application/json"
          },
          body: JSON.stringify(payload)
        });

        if (response.ok) {
          alert("Email forwarded successfully.");
        } else {
          const error = await response.json();
          alert("Failed: " + error.error.message);
        }
      } else {
        alert("Could not get access token.");
      }
    });
  } catch (err) {
    console.error(err);
    alert("Unexpected error occurred.");
  }
}
