# Capstone Project – Aurelia AI Datacenter: guida ai file e passi da fare

Questa cartella contiene **tutti i deliverable richiesti dal professore**, generati da un modello integrato
(schedulazione + procurement + cash-flow) che ho costruito in Python a partire dal workbook del docente.
Tutti i numeri nei vari file sono **coerenti tra loro** perché provengono dallo stesso modello.

## 1. Cosa trovi in `Capstone/deliverables`

| File | Deliverable richiesto | Stato |
|---|---|---|
| `Aurelia_Program_Charter.docx` | Program Charter con Delivery/Release/Contracting Strategy e Procurement & Logistics Plan | Completo (mancano solo nomi del team) |
| `Aurelia_Program_Master_Plan.xml` | Program Master Plan per MS Project / ProjectLibre | Completo, formato XML (da aprire e salvare come `.mpp`, vedi §3) |
| `Aurelia_Program_Risk_Management_Plan.docx` | Risk Management Plan (sintesi Word) | Completo |
| `Aurelia_Risk_Register.xlsx` | Risk Management Plan (dettagli Excel) | Completo, 24 rischi con formule |
| `Aurelia_Financial_Model.xlsx` | Financial Model (FCFF mensile, D/E, WACC, ammortamento, DSCR, NPV) | Completo, ~12.000 formule vive, ricalcolato senza errori |

In `Capstone/model` c'è il codice Python che genera tutto (vedi §6): serve se il docente cambia dati o
se volete esplorare alternative.

## 2. Risultati chiave della soluzione proposta

| Output | Valore |
|---|---|
| NPV massimo fattibile | **260,1 M€** (WACC 11,25 %) |
| Quota debito / equity | **47,5 % / 52,5 %** del CAPEX (D = 285 M€) – massima quota che rispetta DSCR ≥ 1,65x in tutti gli anni |
| CAPEX totale | 600,1 M€ (attività + equipment 580,8; expediting 2,3; trasporti 0,7; overhead 16,3) |
| Hall A in servizio commerciale | **10 maggio 2028** |
| Hall B in servizio commerciale | **2 marzo 2029** |
| Final Program Acceptance | **19 aprile 2029** (deadline 11 settembre 2029 → 145 giorni di margine) |
| DSCR minimo (2028 apr–dic) | 1,66x (vincolo 1,65x); 2029: 3,35x; poi > 3,8x |
| Rimborso | rate mensili uguali apr 2028 – dic 2037, saldo finale 0 |
| Rischi | 24 identificati, 11 High/Critical iniziali, 1 High residuo |

Logica della strategia (spiegata nel Charter, §2 e §9 "Decision log"):
1. Il collo di bottiglia è il **trasformatore di potenza** (50 settimane per lotto, una sola linea produttiva):
   ordine al primo giorno utile (fine F08, 16 apr 2027), **expediting** di entrambi i lotti e trasporto su strada.
2. **Crashing** solo sulle catene guidate dai trasformatori (F23, F24, F29, F31, A08, A10): +0,9 M€ di costo,
   Hall A circa un mese prima.
3. I lotti **compute (160 M€)** e network si pagano all'ordine → ordinati **just-in-time**; per il lotto 1 l'aereo
   (475 k€) costa meno del costo del capitale di ordinare 35 giorni prima.
4. Il test DSCR del 2028 (solo apr–dic, con la sola Hall A) determina quanto debito si può avere: ogni mese di
   ritardo della Hall A toglie ~4 punti di quota debito e ~9 M€ di NPV.

## 3. ATTENZIONE: due fogli del workbook mancano

Il briefing cita i fogli **"Physical Constraints"** e **"Delivery Methods"**, ma il file
`Datacenter_Data_Updated.xlsx` che mi hai dato contiene solo: Read Me, Program Parameters, Financing Inputs,
Activity Dictionary, Resources, Procurement and Logistics, Revenue and Finance.

Ho quindi **ricostruito la logica di precedenza** (quali attività precedono quali, condizioni per installare
l'IT, condizioni di servizio commerciale, regola del corridoio) dalle note delle attività e dal testo del
briefing. Tutte le assunzioni sono elencate nel Charter, §10 (A1–A13) e Annex A (tabella completa dei link).

**Cosa devi fare:**
1. Chiedi al docente / sul portale se esiste una versione del workbook con i due fogli.
2. Se esistono, confronta le regole con l'Annex A del Charter. Se qualche regola è diversa (es. "l'IT può essere
   installato solo dopo il test della hall F41"), modifica la lista `LINKS` in `Capstone/model/engine.py` e
   rigenera tutto (§6). I file si riallineano automaticamente.
3. Per i **Delivery Methods**: ho scelto Design-Build con Early Contractor Involvement per P1, owner-procurement +
   Systems Integrator per P2, in-house + recruitment outsourcing per P3. Verifica che siano tra le opzioni ammesse
   dal foglio e adatta il §5 del Charter se servono nomi/condizioni specifiche.

## 4. Passo-passo: aprire il piano in MS Project e salvarlo come `.mpp`

Io non posso creare direttamente un `.mpp` (formato binario proprietario). L'XML è il formato di scambio
ufficiale di Microsoft Project e si apre senza perdite.

1. Apri **Microsoft Project** (365 / 2019 / 2021, desktop).
2. **File → Apri → Sfoglia**, nel menu a tendina del tipo file scegli **"XML Format (*.xml)"** e seleziona
   `Aurelia_Program_Master_Plan.xml`.
3. Nella finestra "Import Wizard" scegli **"As a new project"** → Fine.
4. Controlla il calendario: **Project → Change Working Time** → calendario "Standard", lun–ven 8:00–12:00 /
   13:00–17:00, nessuna festività. Poi **File → Options → Schedule**: Hours per day 8, Hours per week 40,
   Days per month 20, Default task type **Fixed Duration**, "New tasks are effort driven" **disattivato**.
5. Controlla che tutte le attività siano **Auto Scheduled** (colonna Task Mode). Se alcune risultassero Manually
   Scheduled: selezionale tutte → Task → Auto Schedule.
6. Verifica le date: **Project → Project Information → Statistics**: Finish deve essere **19 apr 2029**;
   milestone "Hall A commercial service" = 9 mag 2028 (ore 17:00; il servizio commerciale parte il giorno dopo,
   10 mag), "Hall B commercial service" = 1 mar 2029, "Final Program Acceptance" = 19 apr 2029.
   Se le date differiscono di qualche giorno, quasi sempre è il calendario (punto 4) o il Task Mode (punto 5).
7. Verifica le risorse: **View → Resource Usage** (o Resource Sheet): nessuna risorsa in rosso
   (nessuna sovra-allocazione). Il piano usa **link espliciti** di sequenza (come nel tuo assignment 1), non il
   livellamento automatico: **non lanciare Resource → Level All**.
8. Verifica i costi: View → Gantt Chart, aggiungi la colonna **Cost**: il totale di progetto deve essere
   ≈ **583,8 M€** (attività 580,8 + expediting 2,3 + trasporti 0,7; gli overhead temporali, 16,3 M€, sono
   solo nel modello finanziario perché il workbook dice che non sono risorse). Le colonne **Text1** = codice
   attività (F01…O15), **Text2** = progetto (P1/P2/P3/Procurement), **Text3** = Train/Hall.
9. Leggi le **Notes** delle attività (icona nella colonna Indicators): contengono le regole di crashing usate, i
   link di sequenza risorse, i costi di expediting e trasporto (inseriti come Fixed Cost).
10. Facoltativo ma consigliato: **Project → Set Baseline** per salvare la baseline, e formatta il Gantt
    (Format → Critical Tasks) per mostrare il percorso critico.
11. **File → Salva con nome → tipo "Project (*.mpp)"** → `Aurelia_Program_Master_Plan.mpp`.
12. Esporta un PDF/immagine del Gantt (File → Print → PDF oppure Format → Copy Picture) e incollalo nel
    Charter, §6 (ho lasciato spazio nel testo).

Alternativa gratuita: **ProjectLibre** → File → Open → tipo "Microsoft Project XML" → stesso file; salva in
`.pod` oppure File → Save as → formato XML.

**Se l'import fallisse** (versione molto vecchia di Project): il foglio `Schedule` del Financial Model
contiene l'elenco completo delle attività con date, durate, risorse e unità (anche quelle di crashing), il
foglio `Procurement` le catene di fornitura e l'Annex A del Charter tutti i link: con quelli il piano si
reinserisce a mano in 1–2 ore.

## 5. Cosa resta da fare a mano (cose che non posso sapere io)

- Nel Charter e nel Risk Plan sostituire **[TEAM LEADER SURNAME]**, i nomi dei membri e la data.
- Inserire nel Charter uno screenshot del Gantt da MS Project (§6).
- Rileggere il Charter §5 (contratti) e §10 (assunzioni) e confermarli con il docente nel Q&A settimanale.
  Domande utili da fare:
  1. "Il CFADS per il DSCR è il flusso operativo (ricavi – elettricità) o include il CAPEX post-aprile 2028?"
     (io ho usato il flusso operativo: con il CAPEX incluso nessun debito passerebbe il test 2028).
  2. "L'installazione IT (A08) può iniziare dopo il fit-out e i test dei train, o serve anche il test della hall
     (F41)?" (se basta il fit-out, Hall A anticipa di ~3 settimane).
  3. "Un ordering event F08 può emettere più PO (uno per lotto)?" (ho assunto di sì, come dice la nota: "one event
     may cover any power packages and lots").
  4. "Sconto mensile: (1+WACC)^(-m/12) oppure (1+WACC/12)^-m?" (nel modello c'è uno switch in `Inputs!B23`).
- Se volete fare la **presentazione**: i numeri sono in `Summary` del Financial Model e nel Charter §2.
- Creare lo ZIP `COGNOMELEADER_CapstoneProject.zip` con i 5 file (XML + .mpp salvato).

## 6. Come rigenerare tutto (se cambiano dati o assunzioni)

Serve Python 3 con `openpyxl`, `python-docx`, `mpxj` (+ Java per MPXJ) e LibreOffice per il ricalcolo.
Dalla cartella `Capstone/model`:

```bash
python3 optimize.py            # cerca la soluzione a NPV massimo (~10 minuti) -> solution.json
```

```bash
python3 build_fin_xlsx.py      # Financial Model Excel
```

```bash
python3 export_msp.py          # XML per MS Project
```

```bash
python3 build_docs.py          # Charter, Risk Plan, Risk Register
```

Dopo `build_fin_xlsx.py` apri il file in Excel (o LibreOffice) e salva per ricalcolare le formule
(io l'ho già fatto per la versione consegnata). Per modificare una regola di precedenza: `engine.py`,
lista `LINKS` (e `NEEDS` per le consegne). Per cambiare una decisione a mano (es. non expeditare un lotto):
modifica `solution.json` e rilancia gli ultimi tre script.

## 7. Struttura del Financial Model (per spiegarlo al docente)

- `Inputs`: dati del caso (blu), leve del team (giallo: convenzione di sconto, mesi di ritardo per sensitività).
- `Schedule`: tutte le attività con date e costi (uguali a MS Project).
- `Procurement`: PO, regola lead-time, expediting, trasporti, consegna vs data richiesta (float).
- `ActivityCost_M`: costo di ogni attività spalmato sui mesi con `NETWORKDAYS` (accrual prorata).
- `Monthly`: CAPEX, overhead, ricavi, elettricità, **FCFF unlevered**, cumulato, fattori di sconto, PV.
- `Financing`: tabella D/E 0–100 % a passi di 0,5 %, WACC, NPV, DSCR per anno, fattibilità, **scelta
  automatica della quota debito massima fattibile**; sensitività covenant 1,50x / 1,80x.
- `Debt_Ledger`: ledger mensile prelievi/interessi/capitale, test DSCR annuale post-grace.
- `Resources`: verifica picchi vs Max Units.
- `Sensitivity`: NPV vs prezzo di vendita, elettricità, WACC.
- `Summary`: tutti gli output richiesti dal briefing.
