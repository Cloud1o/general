# Demo Odoo – Giorgio Scala (Immobiliare Uno / Immobiliare Due)

Database: https://demo1scala.odoo.com — dati demo di fantasia, costruiti sulle
esigenze emerse nella call del 07/10/2026.

## Cosa c'è nel database

| Esigenza del cliente | Dove si mostra in Odoo |
|---|---|
| Due società sulla stessa piattaforma | Multi-company: **Immobiliare Uno S.r.l.** (Napoli) e **Immobiliare Due S.r.l.** (Caserta), piano dei conti IT su entrambe |
| Fatturazione canoni ricorrenti | Abbonamenti: 10 contratti mensili (S00002–S00011), uno per inquilino, riferimento cliente = codice immobile |
| IVA esente art. 10 | Imposta **"0% Art.10 Esente locazioni"** (natura N4, note legali in fattura) su entrambe le società |
| Solleciti automatici con escalation | Contabilità › Solleciti: 5 gg cortese → 20 gg secondo avviso → 30 gg diffida PEC + attività → 60 gg trasmissione al legale (Avv. Martone) |
| Mutui, rate, surroghe | Contabilità › Prestiti: 4 finanziamenti con piano di ammortamento (in bozza) |
| Riconciliazione bancaria | 16 movimenti di ottobre 2026 sui registri Banca (canoni, rate mutuo, condominio, F24) |
| Sostituire Microsoft Planner | Progetti: **Manutenzioni immobili**, **Scadenze fiscali e F24** (flusso Da pagare → Pagato / Avviso bonario → Cartella), **Contenzioso e piani di rientro** |
| Imposta di registro, IMU, F24 | Attività con scadenza nel progetto Scadenze fiscali (una per contratto, IMU acconto/saldo, dichiarazione IMU) |
| Rateizzazione morosi | Ottica Salernitana: nota nel contratto + pratica in "Piano di rientro" |
| Imputazione per immobile | Piano analitico **Immobili** con un conto per unità (IU-01 … ID-06); tag progetto per codice immobile |
| Cash flow medio-lungo termine | Documenti › Gestione immobiliare › Cash flow e budget › **Cash Flow Previsionale 2027** |
| Calcolo IMU | Documenti › Gestione immobiliare › IMU e F24 › **Calcolatore IMU 2027** |
| Rendita % per immobile e patrimonio | Fogli "Immobili" e "KPI Patrimonio" del Cash Flow |

## Prima della demo (5 minuti di clic)

Gli strumenti usati per configurare non possono premere pulsanti di workflow,
quindi questi passaggi vanno fatti a mano:

1. **Abbonamenti** – apri i contratti e premi *Conferma* (oppure lasciane uno in
   bozza e confermalo live in demo). Poi *Crea fattura* su alcuni per avere
   fatture da riconciliare e da sollecitare.
2. **Prestiti** – apri ogni prestito e premi *Conferma* per generare le
   registrazioni.
3. **Distribuzione analitica** – sulle righe dei contratti imposta il conto
   analitico dell'immobile (o crea un modello di distribuzione analitica per
   prodotto/partner).
4. **Documenti** – apri i due fogli di calcolo e controlla che si visualizzino
   correttamente. Se uno non si aprisse, carica il file `.xlsx` corrispondente
   da questa cartella e scegli "Apri con Fogli di calcolo".

## Scaletta demo suggerita (≈45')

1. Cambio società in alto a destra: due P.IVA, un solo database.
2. Contratto di locazione (abbonamento) → fattura esente art. 10 → incasso dalla
   banca → riconciliazione.
3. Inquilino moroso → report solleciti → livelli di escalation fino al legale.
4. Ottica Salernitana: piano di rientro (rateizzazione) e pratica di contenzioso.
5. Progetti come sostituto di Planner: segnalazione manutenzione, preventivo,
   fornitore, tag per immobile.
6. Scadenze fiscali: F24 da pagare, avviso bonario, imposta di registro.
7. Mutui: piano di ammortamento, quota capitale/interessi, simulazione surroga
   (foglio Mutui del Cash Flow).
8. Cash Flow: filtro per società, sfitto (Occupazione), straordinari, proiezione
   5 anni, KPI rendimento lordo/netto e LTV.
9. Calcolatore IMU: beni merce (IMU solo sui mesi locati), acconto/saldo,
   riepilogo F24 per codice tributo, budget vs dichiarazione, aumento aliquota
   Salerno.

## File in questa cartella

- `Cash_Flow_Previsionale_2027.xlsx`, `Calcolatore_IMU_2027.xlsx` – stessi
  contenuti dei fogli in Odoo, apribili anche in Excel.
- `tools/demo_data.py` – dati demo (immobili, mutui, movimenti straordinari).
- `tools/build_spreadsheets.py` – genera i due fogli.

Aliquote IMU e importi sono esempi: vanno verificati sulla delibera del Comune
e con il commercialista.
