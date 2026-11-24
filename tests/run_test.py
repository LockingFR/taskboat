"""Test de bout en bout de Taskboat : sert le site avec les en-têtes de _headers (CSP comprise),
charge une réponse Graph d'exemple (tests/graph-sample.json) et vérifie chaque format d'export.
Usage : pip install playwright && playwright install chromium
        python tests/run_test.py
"""
import asyncio, http.server, json, os, re, sys, threading
from playwright.async_api import async_playwright

HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
HEADERS = [l.strip().split(": ", 1) for l in open(os.path.join(ROOT, "_headers"), encoding="utf-8") if ": " in l]


class H(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **k): super().__init__(*a, directory=ROOT, **k)
    def end_headers(self):
        for k, v in HEADERS:
            if k != "Strict-Transport-Security": self.send_header(k, v)
        super().end_headers()
    def log_message(self, *a): pass


async def export(pg, fmt):
    await pg.click(f'#fmt button[data-f="{fmt}"]')
    return await pg.input_value("#outtext")


async def run():
    errs, checks = [], {}
    async with async_playwright() as p:
        b = await p.chromium.launch(); pg = await b.new_page()
        pg.on("pageerror", lambda e: errs.append(str(e)))
        pg.on("console", lambda m: errs.append(m.text) if m.type == "error" else None)
        # page vue sans identifiant Azure, quel que soit celui du code
        src = open(os.path.join(ROOT, "assets", "app.js"), encoding="utf-8").read()
        js_ct = "text/javascript; charset=utf-8"
        def app(i):
            body = re.sub(r'const CLIENT_ID = "[^"]*"', f'const CLIENT_ID = "{i}"', src)
            return lambda r: r.fulfill(body=body, content_type=js_ct)
        await pg.route("**/assets/app.js", app("NON_CONFIGURE"))
        await pg.goto("http://127.0.0.1:8766/?lang=fr"); await pg.wait_for_timeout(400)
        checks["exemple affiché au chargement"] = "Exemple" in await pg.inner_text("#status")

        await pg.click("#btn-ms"); await pg.wait_for_timeout(200)
        checks["connexion non configurée : message clair"] = "pas encore configurée" in await pg.inner_text("#status")

        await pg.set_input_files("#file", os.path.join(HERE, "graph-sample.json")); await pg.wait_for_timeout(300)
        st = await pg.inner_text("#status")
        checks["JSON Graph chargé (2 listes, 3 tâches)"] = "2 listes, 3 tâches" in st

        csv = await export(pg, "csv")
        checks["CSV : guillemets échappés"] = '"Renvoyer le formulaire, avec ""guillemets"", virgules; et point-virgule"' in csv
        checks["CSV : notes HTML converties"] = "Ligne 2 & suite" in csv

        todo = await export(pg, "todoist")
        checks["Todoist : priorité haute et échéance"] = ",1,1," in todo and "2026-12-05" in todo
        checks["Todoist : rien de terminé"] = "Tâche terminée" not in todo

        md = await export(pg, "md")
        checks["Markdown : échéance, tâche terminée, étapes"] = ("📅 2026-12-05" in md and "- [x] Tâche terminée" in md
                                                               and "    - [x] Éthiopie" in md and "🔁 every 2 weeks" in md)

        ics = await export(pg, "ics")
        checks["ICS : 3 VTODO, DUE, virgules échappées"] = (ics.count("BEGIN:VTODO") == 3 and "DUE;VALUE=DATE:20261205" in ics
                                                           and "avec \"guillemets\"\\, virgules\\;" in ics)

        js = await export(pg, "json")
        data = json.loads(js)
        checks["JSON Taskboat : relu à l'identique"] = data.get("taskboat") == 1 and sum(len(l["tasks"]) for l in data["lists"]) == 3
        os.makedirs(os.path.join(HERE, "out"), exist_ok=True)
        back = os.path.join(HERE, "out", "taskboat-backup.json"); open(back, "w", encoding="utf-8").write(js)
        await pg.set_input_files("#file", back); await pg.wait_for_timeout(300)
        checks["JSON Taskboat : rechargé dans la page"] = "2 listes, 3 tâches" in await pg.inner_text("#status")

        # identifiants Graph longs à préfixe commun : UID tous distincts
        longs = {"value": [{"displayName": "L", "tasks": [{"id": "A" * 120 + str(i), "title": f"t{i}"} for i in range(3)]}]}
        lf = os.path.join(HERE, "out", "longs.json")
        with open(lf, "w", encoding="utf-8") as fh: json.dump(longs, fh)
        await pg.set_input_files("#file", lf); await pg.wait_for_timeout(300)
        unf = re.sub(r"\r?\n ", "", await export(pg, "ics"))
        checks["ICS : UID distincts"] = len(set(re.findall(r"UID:(.+)", unf))) == 3

        steps = {"value": [{"displayName": "S", "tasks": [{"id": "s", "title": "T", "checklistItems": [
            {"displayName": "faite", "isChecked": True}, {"displayName": "reste", "isChecked": False}]}]}]}
        with open(lf, "w", encoding="utf-8") as fh: json.dump(steps, fh)
        await pg.set_input_files("#file", lf); await pg.wait_for_timeout(300)
        todo = await export(pg, "todoist")
        checks["Todoist : étapes cochées exclues"] = "faite" not in todo and "reste,,4,2" in todo

        # connexion simulée : $expand refusé, 429, pagination, date UTC décalée
        ctx = await b.new_context(timezone_id="Europe/Paris"); p2 = await ctx.new_page()
        p2.on("pageerror", lambda e: errs.append(str(e)))
        seen = {"n": 0}
        C = {"Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "Authorization", "Access-Control-Expose-Headers": "Retry-After"}
        async def graph(r):
            u = r.request.url
            async def j(v, status=200, extra=None, **k): await r.fulfill(status=status, headers={**C, **(extra or {})}, json={"value": v, **k})
            if r.request.method == "OPTIONS": return await r.fulfill(status=204, headers=C)
            if "expand" in u: return await j([], 400)
            if "checklistItems" in u:
                seen["n"] += 1
                if seen["n"] == 1: return await j([], 429, {"Retry-After": "1"})
                return await j([{"displayName": "étape", "isChecked": False}])
            if "linkedResources" in u: return await j([])
            if "page2" in u: return await j([{"id": "b", "title": "B"}])
            if "/tasks" in u:
                return await j([{"id": "a", "title": "A", "dueDateTime": {"dateTime": "2026-12-04T23:00:00.0000000", "timeZone": "UTC"}}],
                               **{"@odata.nextLink": "https://graph.microsoft.com/v1.0/tasks?page2"})
            await j([{"id": "l", "displayName": "Liste"}])
        await p2.route("https://graph.microsoft.com/**", graph)
        fake = "00000000-0000-0000-0000-000000000000"
        await p2.route("**/assets/app.js", app(fake))
        await p2.route("**/msal-browser.min.js", lambda r: r.fulfill(content_type=js_ct, body="""window.msal={PublicClientApplication:class{
          getAllAccounts(){return []} async loginPopup(){return {account:{username:"t@example.com"}}} async acquireTokenSilent(){return {accessToken:"x"}}}}"""))
        await p2.goto("http://127.0.0.1:8766/?lang=fr"); await p2.click("#btn-ms"); await p2.wait_for_timeout(3000)
        checks["Graph simulé : sans $expand, 429, pagination"] = "1 liste, 2 tâches" in await p2.inner_text("#status")
        o = json.loads(await export(p2, "json"))["lists"][0]["tasks"]
        checks["Graph simulé : étapes et date locale"] = len(o) == 2 and o[0]["steps"][0]["title"] == "étape" and o[0]["due"] == "2026-12-05"

        # vraie copie locale de MSAL chargée sous la CSP
        p3 = await b.new_page(); p3.on("console", lambda m: errs.append(m.text) if m.type == "error" else None)
        await p3.route("**/assets/app.js", app(fake))
        await p3.goto("http://127.0.0.1:8766/"); await p3.wait_for_timeout(800)
        checks["MSAL local chargé sous la CSP"] = await p3.evaluate("!!(window.msal && msal.PublicClientApplication)")
        await b.close()
    checks["aucune erreur console/CSP"] = not errs
    for k, v in checks.items(): print(("OK   " if v else "ÉCHEC"), k)
    if errs: print("Erreurs :", errs)
    return all(checks.values())


def main():
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 8766), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    ok = asyncio.run(run()); srv.shutdown(); sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
