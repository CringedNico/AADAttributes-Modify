<#
Script Name: ENG-ModificaAADAttributi.ps1
Description: Questo script importa dati da un file Excel e aggiorna le estensioni degli utenti in Azure Active Directory.
Author: Nicolò Bertucci
Date: 12/04/2024
#>

# Importa il modulo necessario per lavorare con file Excel
Import-Module -Name ImportExcel

Invoke-Expression -Command "Connect-AzureAD | Out-null"

# Ottiene i dati dal file di input
$file = $args[0]
$data = Import-Excel -Path $file

# Ottiene la colonna contenente gli indirizzi email
$column = $data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
$mailboxes = $data | Select-Object -ExpandProperty $column[2]

$ids = @()

# Itera attraverso gli indirizzi email e ottiene gli ID corrispondenti
foreach ($mailbox in $mailboxes) {
    $command = "Get-AzureADUser -ObjectId $mailbox | Select-Object -ExpandProperty ObjectId"
    $objectId = Invoke-Expression -Command $command
    $ids += $objectId
}

$employeeids = @()

# Itera attraverso tutte le righe dei dati
foreach ($row in $data) {
    # Ottiene il valore della colonna degli ID degli impiegati
    $employeeid = $row.$($column[1])

    # Controlla se il valore è $null e lo aggiunge alla lista
    if ($employeeid -eq $null) {
        $employeeids += $null
    } else {
        $employeeids += $employeeid
    }
}

$attributi10 = @()

# Itera attraverso tutte le righe dei dati
foreach ($row in $data) {
    # Ottiene il valore della colonna degli attributi 10
    $attributo = $row.$($column[0])

    # Controlla se il valore è $null e lo aggiunge alla lista
    if ($attributo -eq $null) {
        $attributi10 += $null
    } else {
        $attributi10 += $attributo
    }
}

# Determina la lunghezza degli array per ottenere il numero di iterazioni
$arrayLength = $mailboxes.Count

# Itera attraverso gli indici degli array
for ($i = 0; $i -lt $arrayLength; $i++) {
    # Accedi agli elementi degli array con lo stesso indice
    $objectId = $ids[$i]
    $employeeid = $employeeids[$i]
    $attributo10 = $attributi10[$i]

    if ($employeeid -ne $null) {
        Invoke-Expression -Command "Set-AzureADUserExtension -ObjectId $objectId -ExtensionName 'EmployeeID' -ExtensionValue '$employeeid'"
        Write-Host "Set-AzureADUserExtension -ObjectId $objectId -ExtensionName 'EmployeeID' -ExtensionValue '$employeeid'"
    }

    if ($attributo10 -ne $null) {
        Invoke-Expression -Command "Set-AzureADUserExtension -ObjectId $objectId -ExtensionName 'ExtensionAttribute10' -ExtensionValue '$attributo10'"
        Write-Host "Set-AzureADUserExtension -ObjectId $objectId -ExtensionName 'ExtensionAttribute10' -ExtensionValue '$attributo10'"
    }
}