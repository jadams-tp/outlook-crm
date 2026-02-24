Office.onReady(() => {
  document.getElementById("send").onclick = async () => {
    const status = document.getElementById("status");
    status.textContent = "Sending…";

    try {
      const item = Office.context.mailbox.item;

      const subject = item.subject || "";
      const fromName = item.from?.displayName || "";
      const fromEmail = item.from?.emailAddress || "";

      const bodyHtml = await new Promise((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Html, (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
          else reject(res.error);
        });
      });

      const payload = { subject, fromName, fromEmail, bodyHtml };

      const resp = await fetch("https://hook.us2.make.com/t5ed226d56xrg3scp37e5h7l2fn7y8b6", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      if (!resp.ok) throw new Error(`Webhook HTTP ${resp.status}`);

      status.textContent = "✅ Sent to CRM.";
    } catch (e) {
      status.textContent = "❌ Error: " + (e?.message || e);
    }
  };

});
