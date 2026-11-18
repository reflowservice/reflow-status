# REFLOW STATUS

Stato O&M degli impianti fotovoltaici in gestione a Reflow S.r.l.: servizi di manutenzione previsti, uscite dell'anno fatte e da fare, copertura dei servizi, stato contrattuale, mappa degli impianti e scheda di ogni impianto.

Applicativo web (PWA) in sola lettura, con accesso tramite account Microsoft aziendale (registrazione "REFLOW ONE", condivisa con REFLOW PLANNER, REFLOW iMAPS e REFLOW ORBIT). A ogni apertura legge direttamente il file Excel REFLOW MASTER dalla libreria SharePoint aziendale: i dati non stanno in questo repository e non serve nessuno script di aggiornamento.

| File | Ruolo |
|---|---|
| `index.html` | Pagina, stile e logica di calcolo e visualizzazione (`RUN`) |
| `master.js` | Lettura del foglio "Impianti " del Master (blocchi di 8 righe, servizi, stato contrattuale, coordinate) |
| `cloud.js` | Login MSAL e lettura del Master via Microsoft Graph |
| `app.js` | Avvio: accesso, caricamento, schermata iniziale |

## Dati

Nessun dato di impianto, coordinata o nome di cliente va inserito in questo repository: il `.gitignore` esclude i file di dati.

## Manutenzione

Responsabile dell'applicativo: Luca Alberton (PM O&M, Reflow S.r.l.).
