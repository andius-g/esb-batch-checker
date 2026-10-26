import pandas as pd
from playwright.sync_api import sync_playwright

def run_lookup():
    # Excel einlesen (passt sich an, egal ob Name oder Nummer zuerst steht)
    try:
        df = pd.read_excel("spender_liste.xlsx")
    except Exception as e:
        print(f"Fehler beim Laden der Excel: {e}")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        page.goto("https://www.europeanspermbank.com/clinic-de-de/my-page/lookup-tool")
        input("Bitte im Browser einloggen und Enter drücken...")

        results = []

        for index, row in df.iterrows():
            # Nimmt einfach die Werte aus Spalte 0 und 1 und kombiniert sie
            query = f"{str(row.iloc[0]).strip()} {str(row.iloc[1]).strip()}"
            print(f"Prüfe: {query}...")

            try:
                page.goto("https://www.europeanspermbank.com/clinic-de-de/my-page/lookup-tool")
                page.fill('input[placeholder*="Search for donor"]', query)
                page.click('button.btn--rounded')
                
                # Kurzer Warte-Timeout für das Ergebnis
                page.wait_for_timeout(2000)
                
                # Text der Ergebnisseite greifen (Selektor ggf. anpassen)
                content = page.inner_text('.mypage-content-new')
                
                # Ergebnis filtern (Beispielhaft)
                results.append({"Eingabe": query, "Status": "Geprüft", "Details": content[:100]})
            except Exception as ex:
                results.append({"Eingabe": query, "Status": f"Fehler: {ex}", "Details": ""})

        df_out = pd.DataFrame(results)
        df_out.to_excel("ergebnis.xlsx", index=False)
        print("Fertig! Ergebnisse in ergebnis.xlsx gespeichert.")
        browser.close()

if __name__ == "__main__":
    run_lookup()
