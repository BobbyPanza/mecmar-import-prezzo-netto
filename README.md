# Import Prezzo Netto — Excel → Factory

Porta il **Prezzo Netto** del foglio `LISTONE FINALE` del file "Quotazioni rapide"
sulle righe di commessa e sulle parti dell'ERP Factory, per rendere corretto
il calcolo della marginalità.

Specifica funzionale completa: [PRD.md](PRD.md).

---

## Struttura

```
sql/                      oggetti SQL (logica di aggiornamento, modificabile dal consulente)
  01_staging_tables.sql             tabelle di staging e storico
  02_X_fn_trascodeSNpart.sql        trascodifica matricola -> codice parte
  03_X_sp_..._BeginRun.sql          apertura esecuzione / ripartenza
  03b_X_sp_..._GetRows.sql          result set per griglia e report
  04_X_sp_..._Analyze.sql           trascodifica, confronto, calcolo azioni
  05_X_sp_..._Apply.sql             >>> scrittura su ERP <<<
  06_X_sp_..._Rollback.sql          annullamento di un'importazione
  07_backup_pre_apply.sql           fotografia di sicurezza e ripristino integrale
  deploy.ps1                        installa tutto, in ordine, in modo idempotente

src/PrezzoNetto/          applicativo .NET 8 / WPF
  appsettings.json                  configurazione
```

---

## Installazione

### 1. Oggetti SQL

```powershell
cd sql
.\deploy.ps1 -Server "SRV04\Factory" -Database Factory              # autenticazione integrata
.\deploy.ps1 -Server "<ip>\Factory" -Database Factory -User <utente> -Password <password>
```

Gli script sono idempotenti: rilanciarli non perde dati di staging.

### 2. Applicativo

```powershell
cd src\PrezzoNetto
dotnet publish -c Release -r win-x64 --self-contained false -o ..\..\dist
```

Sulla macchina di destinazione serve il **.NET Desktop Runtime 8**.
Per un eseguibile che non richieda runtime installato usare `--self-contained true`.

### 3. Configurazione

`appsettings.json` accanto all'eseguibile:

| Chiave | Significato |
|---|---|
| `ConnectionString` | connessione al DB Factory |
| `Excel.SheetName` | foglio da leggere (default `LISTONE FINALE`) |
| `Excel.HeaderSearchRows` | entro quante righe cercare le intestazioni (nel file reale sono alla riga 2) |
| `Excel.BlankRowsToStop` | righe vuote consecutive che indicano la fine dei dati |
| `Excel.Columns.*` | etichette delle colonne da cercare |
| `Import.Decimals` | decimali di arrotondamento del prezzo |
| `Import.WarnDeltaPct` | scostamento oltre il quale si segnala un warning |
| `Import.ReportFolder` | cartella dei report |

### Credenziali

`appsettings.json` è in controllo di versione e contiene solo valori di esempio.
Le credenziali reali vanno in un **`appsettings.local.json`** accanto
all'eseguibile: sovrascrive le chiavi corrispondenti ed è escluso dal repository.

```json
{
  "ConnectionString": "Server=SRV04\\Factory;Database=Factory;User Id=...;Password=...;TrustServerCertificate=True;Encrypt=False"
}
```

---

## Uso

### Interattivo

Avviare `PrezzoNetto.exe` senza argomenti: **Sfoglia** (o trascinare il file
sulla finestra) → **Analizza** → si controlla la griglia → **Applica**.

Colori della griglia: verde = da aggiornare, giallo = warning, rosso = errore.
Le caselle "Solo differenze" e "Solo problemi" filtrano la vista.
"Simulazione" esegue l'analisi senza scrivere nulla in ERP.

### Da riga di comando

```
PrezzoNetto.exe <file.xlsx> [opzioni]

  --sheet <nome>        foglio da leggere
  --dry-run             analizza senza scrivere su ERP
  --conn <stringa>      stringa di connessione alternativa
  --report <cartella>   cartella del report
  --decimals <n>        decimali di arrotondamento
  --quiet               solo il riepilogo finale
```

| Exit code | |
|---|---|
| 0 | completato senza problemi |
| 1 | completato con warning (righe non trascodificate, ambigue, scostamenti anomali) |
| 2 | errore bloccante |

Esempio di riga per l'Utilità di pianificazione:

```
"C:\Programmi\PrezzoNetto\PrezzoNetto.exe" "\\srv04\Condivisa\Quotazioni rapide.xlsx" --quiet
```

---

## Cosa viene scritto in ERP

| Tabella | Campo | Contenuto |
|---|---|---|
| `L_CMPA` | `PAPRZ` | prezzo netto, sulla riga della commessa cliente (`A_COM.COTYP='P'`) |
| `L_CMPA` | `PriceLastUpdate` | data di applicazione |
| `A_PAR` | `PACSA` | prezzo netto, sulla parte `<MODELLO>-<MATRICOLA>` |

L'applicativo **non** emette `UPDATE` diretti: tutto passa dalle stored procedure.
Per cambiare cosa si scrive si modifica `X_sp_ImportPrezzoNetto_Apply`, senza
ricompilare né ridistribuire l'eseguibile.

---

## Fotografia di sicurezza

Prima di un'importazione importante conviene fotografare i valori attuali di
**tutte** le macchine, non solo di quelle che l'import tocca:

```sql
DECLARE @b int;
EXEC dbo.X_sp_ImportPrezzoNetto_Backup @Label = N'prima import settembre', @BackupId = @b OUTPUT;
```

Lo snapshot copia `L_CMPA.PAPRZ`, `L_CMPA.PriceLastUpdate` e `A_PAR.PACSA` per
ogni parte `<MODELLO>-<MATRICOLA>` e per ogni riga di commessa cliente che la
riferisce (~2.750 righe per tabella). Ogni esecuzione aggiunge uno snapshot
etichettato, nulla viene sovrascritto.

Ripristino integrale:

```sql
EXEC dbo.X_sp_ImportPrezzoNetto_RestoreBackup @BackupId = 1;              -- elenca cosa cambierebbe
EXEC dbo.X_sp_ImportPrezzoNetto_RestoreBackup @BackupId = 1, @WhatIf = 0; -- ripristina davvero
```

## Annullare un'importazione

```sql
-- controllo preventivo: elenca i valori cambiati a mano dopo l'import
EXEC dbo.X_sp_ImportPrezzoNetto_Rollback @RunId = 42;

-- ripristino effettivo
EXEC dbo.X_sp_ImportPrezzoNetto_Rollback @RunId = 42, @Force = 1;
```

Ogni singola scrittura è registrata in `X_ImportPrezzoNetto_Log` con il valore
precedente, quindi ogni esecuzione è reversibile.

---

## Diagnostica

```sql
-- ultime esecuzioni
SELECT TOP 20 * FROM dbo.X_ImportPrezzoNetto_Run ORDER BY RunId DESC;

-- dettaglio di un'esecuzione
EXEC dbo.X_sp_ImportPrezzoNetto_GetRows @RunId = 42;
EXEC dbo.X_sp_ImportPrezzoNetto_GetRows @RunId = 42, @OnlyProblems = 1;

-- verifica della trascodifica di una matricola
SELECT * FROM dbo.X_fn_trascodeSNpart_Detail('226019K');
SELECT dbo.X_fn_trascodeSNpart('226019K');

-- stato dell'ultima importazione applicata, per matricola
SELECT * FROM dbo.X_ImportPrezzoNetto_Current ORDER BY UpdatedAt DESC;
```

### Stato di elaborazione delle righe

| `ProcessStatus` | Rielaborata al rilancio |
|---|---|
| `PENDING` caricata, non ancora analizzata | sì |
| `ANALYZED` pronta da scrivere | sì |
| `DONE` scritta in ERP | no |
| `SKIPPED` nessuna scrittura necessaria | no |
| `ERROR` tentativo fallito | sì |

Per far rielaborare una riga già scritta:

```sql
UPDATE dbo.X_ImportPrezzoNetto_Row
SET ProcessStatus = 'ANALYZED', Applied = 0
WHERE RunId = 42 AND ExcelRow = 57;
```
