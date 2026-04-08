/* global Office */

Office.onReady(async () => {

    // =====================================================
    // 1️⃣ BASE HELPERS
    // =====================================================

    function setStatus(text) {
        const out = document.getElementById("output");
        if (out) out.textContent = text;
    }

    async function callApi(url, payload) {
        const res = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(payload)
        });

        if (!res.ok) {
            const text = await res.text();
            throw new Error(text || res.statusText);
        }

        return res.json();
    }

    function getEmailContext() {
        const item = Office.context.mailbox.item;

        return {
            subject: item.subject || "(No subject)",
            from: item.from?.emailAddress?.address || null,
            sentOn: item.dateTimeCreated || null,
            itemId: item.itemId
        };
    }

    // =====================================================
    // 2️⃣ UI WIRING
    // =====================================================

    // ---------- Log Email ----------
    const logBtn = document.getElementById("logEmail");
    if (logBtn) {
        logBtn.onclick = async () => {
            try {
                setStatus("Logging email…");

                const payload = getEmailContext();
                await callApi("/api/email/log", payload);

                setStatus("✅ Email logged");
            } catch (e) {
                setStatus(`❌ ${e.message}`);
            }
        };
    }

    // ---------- Create Task ----------
    const taskBtn = document.getElementById("createTask");
    if (taskBtn) {
        taskBtn.onclick = async () => {
            try {
                setStatus("Creating task…");

                const payload = getEmailContext();
                await callApi("/api/task/create", payload);

                setStatus("✅ Task created");
            } catch (e) {
                setStatus(`❌ ${e.message}`);
            }
        };
    }

    // ---------- SoOne Lookup ----------
    const searchBtn = document.getElementById("searchSoOne");
    if (searchBtn) {
        searchBtn.onclick = async () => {
            try {
                const entityEl = document.getElementById("entityType");
                const searchEl = document.getElementById("searchText");
                const resultsDiv = document.getElementById("results");

                if (!entityEl || !searchEl || !resultsDiv) {
                    setStatus("Lookup UI not initialised.");
                    return;
                }

                const entity = entityEl.value;
                const searchText = searchEl.value.trim();
                if (!searchText) {
                    setStatus("Enter a search term.");
                    return;
                }

                setStatus("Searching…");

                const results = await callApi("/api/socrm/search", {
                    entity,
                    searchText
                });

                resultsDiv.innerHTML = "";
                results.forEach(r => {
                    const div = document.createElement("div");
                    div.className = "result";
                    div.textContent = `${r.name} (${r.id})`;
                    div.onclick = () => appendIdToSubject(r.id);
                    resultsDiv.appendChild(div);
                });

                setStatus(`Found ${results.length} results`);
            } catch (e) {
                setStatus(`❌ ${e.message}`);
            }
        };
    }

    // ---------- Append ID to Subject ----------
    function appendIdToSubject(id) {
        const item = Office.context.mailbox.item;
        const tag = `[${id}]`;

        if (!item.subject || !item.subject.getAsync) {
            setStatus("Open a new email to tag the subject.");
            return;
        }

        item.subject.getAsync(result => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) return;

            const subject = result.value || "";
            if (subject.includes(tag)) return;

            item.subject.setAsync(`${subject} ${tag}`);
        });
    }
});
