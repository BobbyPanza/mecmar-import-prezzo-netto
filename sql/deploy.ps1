<#
    Deploy degli oggetti SQL dell'import Prezzo Netto sul DB Factory.
    Tutti gli script sono idempotenti: possono essere rieseguiti.

    Esempio:
        .\deploy.ps1 -Server "<ip-o-nome>\Factory" -Database Factory -User <utente> -Password <password>
        .\deploy.ps1 -Server "SRV04\Factory" -Database Factory          # autenticazione integrata
#>
param(
    [string]$Server   = "SRV04\Factory",
    [string]$Database = "Factory",
    [string]$User,
    [string]$Password
)

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

$scripts = @(
    "01_staging_tables.sql",
    "02_X_fn_trascodeSNpart.sql",
    "03_X_sp_ImportPrezzoNetto_BeginRun.sql",
    "03b_X_sp_ImportPrezzoNetto_GetRows.sql",
    "04_X_sp_ImportPrezzoNetto_Analyze.sql",
    "05_X_sp_ImportPrezzoNetto_Apply.sql",
    "06_X_sp_ImportPrezzoNetto_Rollback.sql",
    "07_backup_pre_apply.sql"
)

$auth = if ($User) { @("-U", $User, "-P", $Password) } else { @("-E") }

foreach ($s in $scripts) {
    $path = Join-Path $here $s
    if (-not (Test-Path $path)) { throw "Script non trovato: $path" }
    Write-Host "-> $s" -ForegroundColor Cyan
    & sqlcmd -S $Server -d $Database @auth -b -l 30 -i $path
    if ($LASTEXITCODE -ne 0) { throw "Deploy interrotto su $s (exit $LASTEXITCODE)" }
}

Write-Host "Deploy completato." -ForegroundColor Green
