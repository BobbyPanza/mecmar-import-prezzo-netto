# PRD — Importazione Prezzo Netto da "Listone Finale" in Factory

| | |
|---|---|
| **Versione** | 0.1 (bozza per revisione) |
| **Data** | 2026-09-15 |
| **Autore** | Roberto Russo / Intesi |
| **Cliente / contesto** | Mecmar — ERP Factory (`SRV04\Factory`, DB `Factory`) |
| **Stato** | In definizione — vedere §12 Punti aperti |

---

## 1. Obiettivo

L'ufficio commerciale mantiene i prezzi di vendita effettivi degli essiccatoi in un file Excel
("Quotazioni rapide", foglio `LISTONE FINALE`). L'ERP Factory contiene invece il prezzo di
listino/scontato inserito all'ordine, che spesso **non coincide** con il prezzo netto finale
(trasporti, costi aggiuntivi, provvigioni, trattative successive).

Di conseguenza il calcolo della **marginalità** di commessa in Factory parte da un ricavo errato.

Serve un applicativo che legga il prezzo netto dall'Excel e lo riporti in ERP, sulla riga di
commessa e sull'articolo corrispondenti alla matricola, in modo **ripetibile e riconciliabile**:
ad ogni esecuzione confronta con l'importazione precedente e corregge le differenze.

### Misura di successo
- 100% delle righe con prezzo netto valorizzato viene trascodificato a una riga di commessa
  (sul file campione: 58/58 — vedi §11).
- Ogni scrittura in ERP è tracciata (valore precedente, valore nuovo, run, utente, istante).
- Una ri-esecuzione sullo stesso file non produce scritture (idempotenza).

---

## 2. Ambito

### In ambito
- Applicativo desktop Windows che legge un file Excel e aggiorna il prezzo su Factory.
- Funzione SQL di trascodifica matricola → codice parte.
- Tabelle di staging + storico importazioni.
- Stored procedure T-SQL che, a partire dallo staging, aggiorna l'ERP (modificabile dal consulente).

### Fuori ambito
- Creazione di commesse, righe di commessa o articoli mancanti in ERP.
- Modifica del file Excel (l'applicativo è in sola lettura sul file).
- Ricalcolo di costi, distinte, avanzamenti o report di marginalità (li produce Factory).
- Aggiornamento dei prezzi sulle commesse interne di produzione `CIM*` (`A_COM.COTYP = 'M'`).
- Schedulazione automatica (l'eseguibile è predisposto per essere schedulato, ma la task di
  Windows non è oggetto di fornitura).

---

## 3. Utenti e modalità d'uso

| Utente | Modalità | Scenario |
|---|---|---|
| Ufficio commerciale / controllo di gestione | **Interattiva (WPF)** | Avvio senza argomenti → "Sfoglia" → anteprima differenze → "Applica" |
| Sistemista / task pianificata | **Non presidiata (CLI)** | `PrezzoNetto.exe "\\server\share\Quotazioni.xlsx"` → esecuzione e uscita con exit code |
| Consulente Factory | **SQL** | Modifica la stored procedure di aggiornamento senza toccare l'applicativo |

---

## 4. Il file sorgente

File campione analizzato: `Quotazioni rapide 2023_2 - PROVA 16.07 (1).xlsx`.

### 4.1 Struttura rilevata
- Il workbook contiene 12 fogli; quello di interesse è **`LISTONE FINALE`** (parametrizzabile).
- **La riga di intestazione è la riga 2**, non la 1: la riga 1 contiene i codici opzione
  (`A1_`, `LSTD_`, `P_NETTO`, …). I dati iniziano a riga 3.
- Colonne rilevanti (posizione nel file campione, **da non cablare**):

| Intestazione (riga 2) | Col. | Uso |
|---|---|---|
| `MATRICOLA` | C | Chiave di trascodifica |
| `PREZZO NETTO` | CG | Valore da importare |
| `COMMESSA` | B | Controllo incrociato (vedi §7.5) |
| `CLIENTE` | D | Solo report |
| `MODEL` | E | Controllo incrociato (vedi §7.5) |
| `DATA CONSEGNA` | A | Solo report |

### 4.2 Qualità del dato osservata (116 righe)
| Fenomeno | Esempio | Occorrenze |
|---|---|---|
| Prezzo netto vuoto | righe 3–17 | 58 su 116 |
| Prezzo netto da formula | `=62883-(62883*3/100)` | frequente |
| Prezzo netto digitato a mano | `60996.51` | frequente |
| Matricola numerica | `226035` (numero, non testo) | ~60% |
| Matricola con suffisso lettera | `226019K`, `226043Z`, `226096B` | ~40% |
| Suffisso minuscolo | `226134k`, `226140b`, `226079b` | 4 |
| Riga duplicata | `226139` presente due volte, stesso prezzo | 1 |
| Refuso nella matricola | `2260136Z` (ERP: `226136Z`) | 1 |

**Requisito derivato**: il lettore Excel deve leggere il **valore calcolato** della cella
(non la formula), normalizzare la matricola (trim, maiuscolo, rimozione spazi, conversione
numero→testo senza separatori né decimali) e localizzare le colonne **per nome di intestazione**.

---

## 5. Il modello dati ERP (verificato)

### 5.1 Tabelle
| Tabella | Ruolo |
|---|---|
| `A_COM` | Testata commessa. `COCOD` codice, `CONUM` chiave interna, `COTYP` tipo, `CTCOD`/`CTDSC` cliente |
| `L_CMPA` | Righe di commessa. `CONUM`+`IDROW`, `PACOD` parte, `PAPRZ` prezzo unitario |
| `A_PAR` | Anagrafica parti. `PACOD` codice, `PADSC` descrizione, `PACSA` prezzo |

### 5.2 Codifica della macchina
La macchina finita è una parte codificata **`<MODELLO>-<MATRICOLA>`**, con la matricola anche
in descrizione:

```
PACOD = D20/153THP-226024   PADSC = MECMAR GRAIN DRYER MOD. D20/153T HP SN 226024
PACOD = D24/175T2-226019K   PADSC = MECMAR GRAIN  DRYER MOD. D24/175T2 SN 226019K
```

I componenti di commessa usano invece un prefisso numerico (`226024TESTATA`, `226024CIL-COMP`,
`226024QE`) e **non** matchano il pattern suffisso.

### 5.3 Le due commesse per matricola
Per ogni macchina esistono fino a due commesse:

| Tipo | `COTYP` | Codice | Scopo | `L_CMPA.PAPRZ` |
|---|---|---|---|---|
| Ordine cliente | `P` | `COMM26056` | Vendita | **valorizzato** (es. 65.394) |
| Produzione interna | `M` | `CIM26011` | Costruzione | NULL |

**Il target è la riga della commessa `COTYP = 'P'`.** Sul file campione, tutte le 58 righe con
prezzo netto risolvono a **esattamente una** riga `L_CMPA` di commessa `P` — zero casi mancanti,
zero casi multipli.

### 5.4 Campi da aggiornare (decisi)
| Target | Tipo | Nota |
|---|---|---|
| `L_CMPA.PAPRZ` | `numeric(18,6)` | Riga della commessa `COTYP='P'` |
| `L_CMPA.PriceLastUpdate` | `date` | Data di applicazione dell'import |
| `A_PAR.PACSA` | `numeric(18,6)` | Parte `<MODELLO>-<MATRICOLA>`, `NOT NULL` |

Nota su `PriceLastUpdate`: nel resto del DB è valorizzato su 1.639 righe di `L_CMPA` su 37.518,
sempre in abbinamento a `ListPrice` e mai a `PAPRZ` — l'uso nativo di Factory sembra essere
"data di aggiornamento del prezzo di listino". Sulle 2.743 righe di commessa cliente di
essiccatoi è però `NULL` al 100%, quindi valorizzarlo non sovrascrive alcun dato esistente.

Oggi i due campi risultano allineati tra loro e contengono il prezzo d'ordine
(spesso il *prezzo scontato* dell'Excel, colonna CB, non il *prezzo netto* CG).

### 5.5 Convenzione oggetti custom
Gli oggetti custom nel DB `Factory` usano il prefisso `X_` (`X_IMPORT_TIMESHEET`,
`X_IE_Ordine_Testata`, `X_LogisticsManager_*`). Tutti i nuovi oggetti seguiranno la convenzione.
`X_fn_trascodeSNpart` **non esiste** ed è da creare.

---

## 6. Architettura

```
┌──────────────────────────┐
│  PrezzoNetto.exe (WPF)   │   .NET 8 · ClosedXML · Microsoft.Data.SqlClient
│  ─ Sfoglia / arg CLI     │
│  ─ Lettura + normalizz.  │
│  ─ Griglia anteprima     │
└────────────┬─────────────┘
             │ 1. BULK INSERT staging (una riga per riga Excel)
             │ 2. EXEC X_sp_ImportPrezzoNetto_Analyze  @RunId
             │    ← restituisce l'anteprima delle azioni
             │ 3. EXEC X_sp_ImportPrezzoNetto_Apply    @RunId, @DryRun
             ▼
┌──────────────────────────────────────────────────────────┐
│  DB Factory                                              │
│   X_ImportPrezzoNetto_Run     (storico esecuzioni)       │
│   X_ImportPrezzoNetto_Row     (staging riga per riga)    │
│   X_ImportPrezzoNetto_Current (stato ultima importaz.)   │
│   X_ImportPrezzoNetto_Log     (audit scritture)          │
│   X_fn_trascodeSNpart         (trascodifica SN → parte)  │
│   X_sp_ImportPrezzoNetto_*    (logica modificabile)      │
│        ↓ UPDATE                                          │
│   L_CMPA.PAPRZ · A_PAR.PACSA                             │
└──────────────────────────────────────────────────────────┘
```

**Principio**: l'applicativo non contiene logica di business di aggiornamento. Legge l'Excel,
riempie lo staging, chiama le stored procedure e mostra il risultato. Tutta la logica di
trascodifica, confronto e aggiornamento vive in T-SQL ed è modificabile dal consulente Factory
senza ricompilare né ridistribuire l'eseguibile.

### 6.1 Stack
| Componente | Scelta |
|---|---|
| Runtime | .NET 8, `win-x64`, self-contained single-file |
| UI | WPF |
| Excel | ClosedXML (nessuna dipendenza da Office installato) |
| Accesso dati | `Microsoft.Data.SqlClient` (+ Dapper per la lettura degli esiti) |
| Configurazione | `appsettings.json` accanto all'eseguibile |
| Log | Serilog su file rotativo giornaliero + tabella `X_ImportPrezzoNetto_Log` |

---

## 7. Requisiti funzionali

### RF-1 — Selezione del file
- **RF-1.1** Avvio senza argomenti: finestra WPF con pulsante **Sfoglia** (`OpenFileDialog`,
  filtro `*.xlsx;*.xlsm`), campo percorso editabile, drag&drop del file sulla finestra.
- **RF-1.2** Avvio con percorso come primo argomento: nessuna finestra, esecuzione e uscita.
- **RF-1.3** Percorsi UNC supportati. Se il file è aperto in Excel da un altro utente,
  va comunque letto (apertura in sola lettura, senza lock).
- **RF-1.4** L'ultimo percorso usato viene ricordato e riproposto.

### RF-2 — Parametrizzazione
Tutti i seguenti valori sono in `appsettings.json` e sovrascrivibili da riga di comando:

| Parametro | Default | CLI |
|---|---|---|
| Stringa di connessione | `Server=SRV04\Factory;Database=Factory;…` | `--conn` |
| Nome foglio | `LISTONE FINALE` | `--sheet` |
| Etichetta colonna matricola | `MATRICOLA` | `--col-matricola` |
| Etichetta colonna prezzo | `PREZZO NETTO` | `--col-prezzo` |
| Righe in cui cercare l'intestazione | `1..10` | `--header-rows` |
| Righe vuote consecutive = fine dati | `20` | — |
| Decimali di arrotondamento | `2` | `--decimals` |
| Modalità | `apply` | `--dry-run` |
| Cartella report | `.\Report` | `--report` |

- **RF-2.1** Il nome del foglio è confrontato senza distinzione di maiuscole/minuscole e
  ignorando spazi iniziali/finali.
- **RF-2.2** Le intestazioni sono cercate nelle prime N righe del foglio; il confronto è
  normalizzato (maiuscolo, trim, spazi multipli compattati). Se una delle due colonne non
  viene trovata, l'esecuzione termina con errore esplicito che elenca le intestazioni trovate.

### RF-3 — Lettura e normalizzazione
- **RF-3.1** Delle celle con formula si legge il **valore memorizzato**; se il valore non è
  disponibile (workbook mai ricalcolato) la riga è marcata `ERRORE_VALORE_NON_DISPONIBILE`.
- **RF-3.2** Normalizzazione matricola: `TRIM` → maiuscolo → rimozione spazi e caratteri non
  alfanumerici → se numerica, formato intero senza decimali (`226035.0` → `226035`).
- **RF-3.3** Riga senza matricola: ignorata silenziosamente (non conteggiata come errore).
- **RF-3.4** Riga con matricola ma **senza** prezzo netto (o prezzo = 0, o non numerico):
  registrata in staging con azione `SKIPPED_NOPRICE`, **nessuna scrittura**. Non è un errore:
  sul file campione sono 58 righe su 116.
- **RF-3.5** Matricola ripetuta nello stesso file:
  - stesso prezzo → si consolida in un'unica riga logica (caso `226139`);
  - prezzi diversi → **entrambe** le righe marcate `AMBIGUO_EXCEL`, nessuna scrittura.
- **RF-3.6** Il prezzo netto è arrotondato a 2 decimali (configurabile) prima del confronto
  e della scrittura.

### RF-4 — Trascodifica matricola → parte
Realizzata dalla funzione SQL `X_fn_trascodeSNpart` (§8).

- **RF-4.1** Livello 1 — **esatto**: esiste una parte il cui `PACOD` termina con `-<matricola>`.
- **RF-4.2** Livello 2 — **nucleo numerico**: nessun match esatto → si cerca `-<nucleo>` o
  `-<nucleo><lettera>`, dove *nucleo* sono le cifre iniziali della matricola. Copre i casi reali
  Excel `226012` ↔ ERP `226012K` ed Excel `226034B` ↔ ERP `…-226034`.
  **Non** copre il refuso `2260136Z` (nucleo `2260136`, sette cifre, contro `226136` in ERP):
  è corretto che resti non trascodificato anziché agganciarsi alla macchina sbagliata.
- **RF-4.3** In caso di candidati multipli, il tie-break è, nell'ordine:
  1. la parte che ha una riga in una commessa `COTYP='P'` (la più recente per `CONUM`);
  2. la parte non bloccata (`IsBlocked = 0`);
  3. la parte con `IDART` più alto (la più recente).

  Verificato sui due casi reali del file campione (`226022K`, `226068K`): il criterio 1 li
  risolve entrambi univocamente.
- **RF-4.4** Nessun candidato → riga marcata `NON_TRASCODIFICATA`, nessuna scrittura, warning
  a report. Il match di livello 2 viene comunque **segnalato** come warning in anteprima.

### RF-5 — Risoluzione riga di commessa
- **RF-5.1** Dalla parte si risale alla riga `L_CMPA` collegata a una testata `A_COM` con
  `COTYP = 'P'`.
- **RF-5.2** Zero righe → `COMMESSA_NON_TROVATA`, nessuna scrittura.
- **RF-5.3** Più righe → si sceglie la commessa con `CONUM` più alto e la riga viene marcata
  `COMMESSA_MULTIPLA` (warning, ma la scrittura procede). *(Da confermare — §12)*

### RF-6 — Confronto con l'ultima importazione
- **RF-6.1** La tabella `X_ImportPrezzoNetto_Current` conserva, per matricola, il prezzo netto
  applicato l'ultima volta e i valori effettivamente scritti su `PAPRZ` e `PACSA`.
- **RF-6.2** Azioni calcolate per ogni riga:

| Azione | Condizione |
|---|---|
| `NUOVO` | matricola mai importata |
| `MODIFICATO` | prezzo Excel ≠ prezzo dell'ultima importazione |
| `INVARIATO` | prezzo Excel = ultima importazione **e** valori ERP invariati → nessuna scrittura |
| `RIALLINEATO` | prezzo Excel invariato ma il valore in ERP è stato cambiato a mano → si riscrive |
| `SKIPPED_NOPRICE` / `NON_TRASCODIFICATA` / `AMBIGUO_EXCEL` / `COMMESSA_NON_TROVATA` | vedi sopra |

- **RF-6.3** **Modifiche manuali in ERP**: il valore Excel è la fonte di verità e viene sempre
  sovrascritto; il valore precedente e il flag "modifica manuale rilevata" finiscono in
  `X_ImportPrezzoNetto_Log`, così la modifica resta tracciata e recuperabile.
- **RF-6.4** Matricola presente nelle importazioni precedenti e **sparita** dall'Excel:
  nessuna azione su ERP, segnalazione in report come `ASSENTE_IN_SORGENTE`.

### RF-7 — Applicazione
- **RF-7.1** L'aggiornamento è eseguito dalla stored procedure `X_sp_ImportPrezzoNetto_Apply`
  in **un'unica transazione**: o passa tutto o non passa niente.
- **RF-7.2** Ogni `UPDATE` scrive una riga in `X_ImportPrezzoNetto_Log` con valore precedente,
  valore nuovo, tabella/campo, chiave, run, utente, timestamp.
- **RF-7.3** Modalità **dry-run**: eseguono analisi, staging e anteprima, ma nessun `UPDATE`.
- **RF-7.4** In modalità interattiva l'utente vede l'anteprima **prima** di applicare e conferma
  esplicitamente.
- **RF-7.5** **Ripartenza**: se l'applicazione si interrompe, un nuovo avvio rileva il run
  rimasto in stato `RUNNING` sullo stesso file (confronto per `SourceFileHash`) e propone di
  **riprenderlo** anziché ricominciare: vengono rielaborate solo le righe non ancora `DONE`
  (§9.2.1).

### RF-8 — Interfaccia WPF
- **RF-8.1** Schermata unica: percorso file + Sfoglia, nome foglio, casella "Simulazione",
  pulsante **Analizza**, griglia risultati, pulsante **Applica** (abilitato solo dopo analisi).
- **RF-8.2** Colonne griglia: riga Excel, Matricola, Modello, Cliente, Commessa Excel,
  Parte ERP, Commessa ERP, Prezzo ERP attuale, Prezzo netto Excel, Delta, Azione, Messaggio.
- **RF-8.3** Colorazione per azione (nuovo / modificato / invariato / warning / errore) e
  filtri rapidi "solo differenze", "solo problemi".
- **RF-8.4** Barra di riepilogo: righe lette, con prezzo, trascodificate, da aggiornare,
  invariate, in errore.
- **RF-8.5** Esportazione del risultato in `.xlsx` nella cartella report.
- **RF-8.6** Le righe in errore non bloccano l'applicazione delle righe corrette.

### RF-9 — Interfaccia a riga di comando
```
PrezzoNetto.exe <percorso-file> [opzioni]

  --sheet <nome>        Foglio da leggere (default: LISTONE FINALE)
  --dry-run             Analizza e produce il report senza scrivere su ERP
  --conn <stringa>      Stringa di connessione alternativa
  --report <cartella>   Cartella di destinazione del report
  --quiet               Nessun output a video oltre al riepilogo finale
```

Exit code:

| Codice | Significato |
|---|---|
| `0` | Completato, nessun problema |
| `1` | Completato con warning (righe non trascodificate, ambigue, commesse multiple) |
| `2` | Errore bloccante (file/foglio/colonne non trovati, DB non raggiungibile, transazione fallita) |

---

## 8. Funzione di trascodifica `X_fn_trascodeSNpart`

### 8.1 Firma
```sql
-- Motore: restituisce anche il livello di match e il numero di candidati
CREATE FUNCTION dbo.X_fn_trascodeSNpart_Detail (@SN varchar(30))
RETURNS TABLE
-- colonne: PACOD varchar(20), MatchLevel varchar(10), Candidates int, IDART int

-- Wrapper scalare per uso diretto in query e viste
CREATE FUNCTION dbo.X_fn_trascodeSNpart (@SN varchar(30))
RETURNS varchar(20)
```

### 8.2 Algoritmo
1. **Normalizzazione**: `UPPER(LTRIM(RTRIM(@SN)))`, rimozione di spazi e caratteri non
   alfanumerici → `@norm`.
2. **Scomposizione**: `@core` = sole cifre di `@norm`; `@suffix` = lettere finali.
3. **Livello `EXACT`**: parti con `RIGHT(RTRIM(PACOD), LEN(@norm)+1) = '-' + @norm`.
4. **Livello `CORE`** (solo se il livello 1 non produce candidati): parti il cui codice termina
   con `-@core` oppure con `-@core` + una sola lettera.
5. **Tie-break** (RF-4.3): riga di commessa `COTYP='P'` più recente → non bloccata → `IDART` max.
6. Nessun candidato → `NULL`, `MatchLevel = 'NONE'`.

### 8.3 Note implementative
- Il pattern `-<matricola>` in coda al codice **non** intercetta i componenti di commessa
  (`226024CIL-COMP`, `226024QE`): verificato sui 2.754 codici che rispettano il pattern
  `…-<6 cifre>[lettera]`, tutti macchine finite.
- Il confronto **non** deve basarsi su `PADSC`: le descrizioni alternano `DRYER` e `DRIER`,
  `SN` e `S.N.`, con spaziature irregolari.
- Il confronto **non** deve filtrare su `PATYP`: le macchine sono sia `F` che `S`.
- **Collation**: il DB `Factory` è uniformemente `Latin1_General_CI_AS` (5.948 colonne utente su
  5.949); l'unica eccezione è `dbo.A_STA.STTYP` in `Latin1_General_CS_AS`, tabella che non
  tocchiamo. Su `A_PAR`, `L_CMPA` e `A_COM` non serve alcun `COLLATE`. Va invece messo
  `COLLATE DATABASE_DEFAULT` nelle sole query che interrogano il catalogo di sistema
  (`sys.*`, `INFORMATION_SCHEMA`), le cui colonne usano la collation fissa dei metadati
  `Latin1_General_CI_AS_KS_WS`: senza, si prende l'errore 451.
- I campi `PACOD`/`CORIF` contengono spazi di riempimento: usare sempre `RTRIM`/`LTRIM`.

---

## 9. Modello dati di staging

Tutte le tabelle nel DB `Factory`, schema `dbo`, prefisso `X_ImportPrezzoNetto_`.

### 9.1 `X_ImportPrezzoNetto_Run` — storico esecuzioni
| Colonna | Tipo | Note |
|---|---|---|
| `RunId` | `int IDENTITY` PK | |
| `StartedAt` / `FinishedAt` | `datetime2` | |
| `SourceFile` | `nvarchar(500)` | percorso completo |
| `SourceFileHash` | `char(64)` | SHA-256: riconosce la ri-esecuzione sullo stesso file |
| `SourceFileModifiedAt` | `datetime2` | |
| `SheetName` | `nvarchar(128)` | |
| `Mode` | `varchar(10)` | `APPLY` / `DRYRUN` |
| `Status` | `varchar(20)` | `RUNNING` / `COMPLETED` / `FAILED` |
| `UserName` / `MachineName` | `nvarchar(128)` | |
| `RowsRead`, `RowsWithPrice`, `RowsMatched`, `RowsUpdated`, `RowsUnchanged`, `RowsWarning`, `RowsError` | `int` | |
| `ErrorMessage` | `nvarchar(max)` | |

### 9.2 `X_ImportPrezzoNetto_Row` — staging
| Colonna | Tipo | Note |
|---|---|---|
| `RunId`, `ExcelRow` | `int` | PK composita |
| `MatricolaRaw` | `nvarchar(50)` | come letta |
| `Matricola` | `varchar(30)` | normalizzata |
| `PrezzoNetto` | `numeric(18,6)` | arrotondato |
| `CommessaExcel` | `int NULL` | colonna `COMMESSA` |
| `Modello`, `Cliente` | `nvarchar(160)` | contesto/report |
| `DataConsegna` | `date NULL` | |
| `PACOD` | `varchar(20) NULL` | esito trascodifica |
| `MatchLevel` | `varchar(10)` | `EXACT` / `CORE` / `NONE` |
| `CONUM`, `IDROW` | `int NULL` | riga di commessa risolta |
| `COCOD` | `varchar(15) NULL` | |
| `OldPAPRZ`, `OldPACSA` | `numeric(18,6) NULL` | valori ERP prima della scrittura |
| `PrevRunPrezzo` | `numeric(18,6) NULL` | prezzo dell'ultima importazione |
| `Action` | `varchar(30)` | vedi RF-6.2 |
| `ProcessStatus` | `varchar(15)` | **stato di elaborazione della riga**, vedi §9.2.1 |
| `ProcessedAt` | `datetime2 NULL` | istante in cui la riga è passata a `DONE`/`SKIPPED`/`ERROR` |
| `ProcessAttempts` | `int NOT NULL DEFAULT 0` | numero di tentativi di elaborazione |
| `Message` | `nvarchar(500)` | esito/errore dell'ultimo tentativo |

#### 9.2.1 Stato di elaborazione della riga
Ogni riga di staging porta il proprio stato, così che `X_sp_ImportPrezzoNetto_Apply` possa essere
rilanciata senza rielaborare ciò che ha già fatto:

| `ProcessStatus` | Significato | Rielaborata da un nuovo `Apply`? |
|---|---|---|
| `PENDING` | caricata dall'Excel, non ancora analizzata | sì |
| `ANALYZED` | trascodificata e confrontata, in attesa di conferma | sì |
| `DONE` | scritta su ERP con successo | **no** |
| `SKIPPED` | nessuna scrittura necessaria (`INVARIATO`, `SKIPPED_NOPRICE`) | no |
| `ERROR` | tentativo fallito (riga di commessa sparita, violazione di vincolo, …) | sì, al rilancio |

Conseguenze operative:
- **Idempotenza**: se l'applicazione si interrompe a metà (rete, riavvio, timeout), il rilancio
  dello stesso run riprende dalle sole righe `PENDING` / `ANALYZED` / `ERROR`.
- **Ripartenza selettiva**: il consulente può rimettere a mano una riga da `DONE` a `ANALYZED`
  per farla rielaborare, senza rifare l'intero import.
- **Diagnosi**: `ProcessAttempts` + `Message` mostrano subito quali righe hanno faticato e perché.
- L'indice `IX_Row_Pending (RunId, ProcessStatus) INCLUDE (Action)` serve la selezione delle
  righe da elaborare.

Le righe dello staging **non** vengono cancellate: sono lo storico. Retention configurabile
(default: mantenere tutto; vedi §12).

### 9.3 `X_ImportPrezzoNetto_Current` — stato ultima importazione applicata
| Colonna | Tipo |
|---|---|
| `Matricola` | `varchar(30)` PK |
| `PACOD`, `CONUM`, `IDROW`, `COCOD` | |
| `PrezzoNetto` | `numeric(18,6)` |
| `WrittenPAPRZ`, `WrittenPACSA` | `numeric(18,6)` — per rilevare modifiche manuali |
| `SourceRunId`, `UpdatedAt`, `UpdatedBy` | |

### 9.4 `X_ImportPrezzoNetto_Log` — audit delle scritture
| Colonna | Tipo |
|---|---|
| `LogId` | `bigint IDENTITY` PK |
| `RunId`, `Matricola` | |
| `TargetTable` | `varchar(30)` — `L_CMPA` / `A_PAR` |
| `TargetColumn` | `varchar(30)` — `PAPRZ` / `PACSA` |
| `TargetKey` | `nvarchar(100)` — es. `CONUM=13869;IDROW=110358` |
| `OldValue`, `NewValue` | `numeric(18,6)` |
| `ManualChangeDetected` | `bit` |
| `WrittenAt`, `WrittenBy` | |

---

## 10. Stored procedure (logica modificabile dal consulente)

| Procedura | Responsabilità |
|---|---|
| `X_sp_ImportPrezzoNetto_BeginRun` | Crea la riga di `_Run`, restituisce `@RunId` |
| `X_sp_ImportPrezzoNetto_Analyze` | Elabora le sole righe `ProcessStatus = 'PENDING'`: trascodifica, risolve la riga di commessa, legge i valori ERP correnti, confronta con `_Current`, calcola `Action` e `Message`, porta la riga ad `ANALYZED` (o `SKIPPED`/`ERROR`). Nessuna scrittura su ERP. Restituisce il result set per la griglia |
| `X_sp_ImportPrezzoNetto_Apply` | `@RunId`, `@DryRun bit`. In transazione, sulle sole righe `ProcessStatus IN ('ANALYZED','ERROR')` con `Action IN ('NUOVO','MODIFICATO','RIALLINEATO')`: `UPDATE L_CMPA.PAPRZ` + `PriceLastUpdate`, `UPDATE A_PAR.PACSA`; scrive `_Log`; aggiorna `_Current`; porta la riga a `DONE` (o `ERROR`, incrementando `ProcessAttempts`); chiude `_Run`. Rilanciabile: le righe già `DONE` vengono ignorate |
| `X_sp_ImportPrezzoNetto_Rollback` | `@RunId`. Ripristina da `_Log` i valori precedenti di un run (rete di sicurezza) |

**Vincolo**: l'applicativo non emette alcun `UPDATE` diretto su `L_CMPA` o `A_PAR`. Tutte le
regole di scrittura (quali campi, quali filtri, eventuali campi aggiuntivi come date di ultimo
aggiornamento prezzo) sono modificabili intervenendo solo su queste procedure.

---

## 11. Verifica di fattibilità (già eseguita sul file campione)

| Verifica | Esito |
|---|---|
| Righe dati nel foglio | 116 |
| Righe con prezzo netto valorizzato | 58 |
| Matricole trascodificate a una parte (match esatto) | 112 / 116 |
| Matricole risolte solo col nucleo numerico | 2 (`226012`, `226034B`) |
| Matricole non trascodificabili | 1 (`2260136Z`, refuso nell'Excel) |
| Matricole con 2 parti candidate | 2 (`226022K`, `226068K`) — risolte dal tie-break |
| **Righe con prezzo netto → 1 sola riga di commessa `COTYP='P'`** | **58 / 58** |
| Righe con prezzo ERP ≠ prezzo netto Excel (da correggere) | ~40 su 58 |
| Corrispondenza colonna `COMMESSA` Excel ↔ `COCOD` ERP | confermata (226024 → 56 → `COMM26056`) |

Nessuna delle 4 matricole problematiche ha oggi un prezzo netto valorizzato: **l'impatto
attuale è nullo**, ma il livello di match `CORE` serve a coprirle quando lo avranno.

### Anomalie di qualità dato rilevate (da segnalare, non da correggere automaticamente)
- `226159`: Excel indica modello `NANO5/40T`, in ERP la parte è `FSN12/110F-226159`.
- `226095Z`: prezzo ERP 96.182,20 contro netto Excel 117.291,65 (scostamento del 22%).
- `226045B`: prezzo ERP 5.000 contro netto Excel 31.000.
- `226159`: prezzo ERP 32.137 contro netto Excel 80.189,21.

### 11.1 Controlli incrociati di coerenza
Non bloccano la scrittura, ma generano warning in anteprima e a report:
- **Commessa**: il numero in colonna `COMMESSA` dell'Excel deve corrispondere alla parte
  numerica del `COCOD` ERP.
- **Modello**: il `MODEL` dell'Excel, normalizzato (spazi rimossi, maiuscolo), deve essere
  prefisso del `PACOD` ERP.
- **Scostamento**: variazione del prezzo oltre una soglia configurabile (default 30%)
  rispetto al valore ERP attuale.

---

## 12. Punti aperti / decisioni da confermare

| # | Domanda | Proposta |
|---|---|---|
| 1 | **Arrotondamento**: il netto Excel ha fino a 4 decimali (`80189,2064`). A quanti decimali scriviamo? | 2 decimali |
| 2 | **Comportamento di default da CLI**: passando solo il percorso, applica o simula? | Applica; `--dry-run` per simulare |
| 3 | ~~`L_CMPA.PriceLastUpdate` e `ListPrice` vanno valorizzati?~~ | **Deciso**: si valorizza `PriceLastUpdate` con la data di applicazione; `ListPrice` non si tocca |
| 4 | **Login SQL**: usiamo `sa`? | Creare un login dedicato con permessi minimi (`SELECT` su `A_COM`/`A_PAR`/`L_CMPA`, `UPDATE` sui soli due campi target, `EXEC` sulle procedure) |
| 5 | **Riga di commessa multipla** (RF-5.3): scrivere sulla più recente o bloccare? | Sul campione non si verifica mai; proposta: scrivere e segnalare |
| 6 | **Retention** dello staging | Mantenere tutto; valutare purge oltre 24 mesi |
| 7 | La `COMMESSA` dell'Excel va usata come chiave di **fallback** quando la matricola non trascodifica? | Da valutare in fase 2 |
| 8 | Notifica via e-mail del report a fine esecuzione non presidiata | Fuori ambito in v1 |

---

## 13. Piano di rilascio

| Fase | Contenuto | Verifica |
|---|---|---|
| **1 — SQL** | `X_fn_trascodeSNpart(_Detail)`, tabelle di staging, 4 stored procedure | Trascodifica su tutte le matricole storiche di `A_PAR`; confronto con l'analisi già eseguita |
| **2 — Lettore Excel** | Apertura, ricerca foglio/intestazioni, normalizzazione, caricamento staging | 116 righe lette dal file campione con gli stessi esiti del §11 |
| **3 — UI WPF** | Sfoglia, griglia, filtri, applica, export | Test manuale |
| **4 — CLI** | Argomenti, exit code, log su file | Esecuzione da task pianificata |
| **5 — Collaudo** | Esecuzione in dry-run su copia del DB di produzione, revisione del report con il commerciale | Report approvato prima del primo `APPLY` |

### Esito del collaudo (15/09/2026)

Le fasi 1–5 sono state completate direttamente su `SRV04\Factory`, previa
fotografia di sicurezza (`X_sp_ImportPrezzoNetto_Backup`, `BackupId = 1`:
2.743 righe di commessa e 2.754 parti).

| Verifica | Esito |
|---|---|
| Funzione di trascodifica su casi noti | 13 / 13 |
| Lettura del file campione | 116 righe, intestazioni riconosciute alla riga 2 |
| Righe con prezzo netto | 58, di cui 1 duplicata e consolidata → 57 da scrivere |
| Applicazione | 57 righe, run `COMPLETED` |
| Righe di log prodotte | 171 = 57 × 3 campi |
| `L_CMPA.PAPRZ`, `A_PAR.PACSA`, `PriceLastUpdate` allineati al prezzo netto | 57 / 57, 0 disallineamenti |
| Righe effettivamente modificate rispetto allo snapshot | 46 (le altre 11 avevano già il valore corretto) |
| **Idempotenza**: seconda esecuzione sullo stesso file | 0 righe da aggiornare, 57 `INVARIATO`, 0 scritture |

Warning residui: 11 alla prima esecuzione, 10 alla seconda — tutti da rivedere
con il commerciale, non difetti dell'importazione (vedi §11).
