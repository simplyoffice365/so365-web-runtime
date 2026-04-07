/* global Office */

Office.onReady(async () => {

    // =====================================================
    // 1️⃣ BASE HELPERS
    // =====================================================

    const SITE_URL = "https://simplyoffice365.sharepoint.com/sites/MW12";

    function setStatus(text) {
        const out = document.getElementById("output");
        if (out) out.textContent = text;
    }

    async function getSsoToken() {
        return await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true
        });
    }

    function spHeaders(token) {
        return {
            "Authorization": `Bearer ${token}`,
            "Accept": "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata"
        };
    }

    // =====================================================
    // 2️⃣ SHAREPOINT REST FUNCTIONS
    // =====================================================

    async function createEmailActivity(token, subject) {
        const url =
            `${SITE_URL}/_api/web/lists/getbytitle('Activity Email')/items`;

        const payload = {
            Title: subject,
            EmailBody: "Logged from Outlook add-in"
        };

        const res = await fetch(url, {
            method: "POST",
            headers: spHeaders(token),
            body: JSON.stringify(payload)
        });

        if (!res.ok) throw new Error(await res.text());
    }

    async function createTaskActivity(token, subject) {
        const url =
            `${SITE_URL}/_api/web/lists/getbytitle('Activity Task')/items`;

        const payload = {
            Title: subject,
            Description: "Task created from Outlook email"
        };

        const res = await fetch(url, {
            method: "POST",
            headers: spHeaders(token),
            body: JSON.stringify(payload)
        });

        if (!res.ok) throw new Error(await res.text());
    }

    async function searchList(token, listName, searchText) {
        const filter = encodeURIComponent(
            `substringof('${searchText}', Title)`
        );

        const url =
            `${SITE_URL}/_api/web/lists/getbytitle('${listName}')/items` +
            `?$select=Id,Title&$filter=${filter}&$top=10`;

        const res = await fetch(url, {
            headers: spHeaders(token)
        });

        if (!res.ok) throw new Error(await res.text());

        return (await res.json()).value;
    }

    async function searchSoCRM(token, entity, searchText) {
        const map = {
            AC: "Accounts",
            CO: "Contacts",
            SL: "Sales Leads",
            SO: "Sales Opportunities",
            PR: "Projects",
            CR: "Cases"
        };

        const listName = map[entity];
        if (!listName) return [];

        const items = await searchList(token, listName, searchText);

        return items.map(i => ({
            id: `${entity}-${i.Id}`,
            name: i.Title
        }));
    }

    // =====================================================
    // 3️⃣ UI WIRING
    // =====================================================

    // ---------- Log Email ----------
    const logBtn = document.getElementById("logEmail");
    if (logBtn) {
        logBtn.onclick = async () => {
            try {
                setStatus("Logging email...");
                const token = await getSsoToken();

                const item = Office.context.mailbox.item;
                const subject = item.subject || "(No subject)";

                await createEmailActivity(token, subject);
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
                setStatus("Creating task...");
                const token = await getSsoToken();

                const subject = Office.context.mailbox.item.subject || "(No subject)";
                await createTaskActivity(token, subject);

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

                setStatus("Searching...");
                const token = await getSsoToken();
                const results = await searchSoCRM(token, entity, searchText);

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