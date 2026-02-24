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

      const resp = await fetch("MAKE_WEBHOOK_URL", {
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