# Spend Analyser - Multi-Account Statement Consolidator & Deduplicator
# Combines:
# 1. PhonePe Statement (Primary UPI Ledger)
# 2. ICICI Credit Card Statement (Direct Card Swipes)
# 3. HDFC Bank Statement (ATM Cash, Inflows, Unique Direct Debits, CC Bill Repayments)
# 4. SBI Bank Statement (Appended if unencrypted file is available)

param(
  [string]$MonthFolder = "C:\Abhishek\August-2026",
  [string]$OutputFile = "C:\Abhishek\August-2026\August-Expenditure - 2026.xlsx"
)

Write-Output "Starting statement consolidation for: $MonthFolder"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  $allTxns = [System.Collections.ArrayList]::new()
  $utrSet = New-Object System.Collections.Generic.HashSet[string]

  # Load rules from rules.json
  $rulesFile = Join-Path $PSScriptRoot "rules.json"
  if (-not (Test-Path $rulesFile)) { $rulesFile = ".\rules.json" }
  $rules = @()
  if (Test-Path $rulesFile) {
    Write-Output "Loading classification rules from: $rulesFile"
    $rules = Get-Content $rulesFile -Raw | ConvertFrom-Json
    Write-Output "Loaded $($rules.Count) rules from rules.json"
  }

  function Get-CategoryFromRules($text, $defaultDomain = "Misc") {
    $low = $text.ToLower()
    foreach ($r in $rules) {
      if ($low.Contains($r.keyword.ToLower())) {
        return $r.domain
      }
    }
    return $defaultDomain
  }

  # Find statement files in folder
  $ppFile = Get-ChildItem -Path $MonthFolder -Filter "*PhonePe*.xlsx" | Select-Object -First 1
  if (-not $ppFile) { $ppFile = Get-ChildItem -Path $MonthFolder -Filter "*PhonePe*.csv" | Select-Object -First 1 }
  $icFile = Get-ChildItem -Path $MonthFolder -Filter "*ICICI*.xlsx" | Select-Object -First 1
  $hdFile = Get-ChildItem -Path $MonthFolder -Filter "*HDFC*.xls*" | Select-Object -First 1
  $sbiFile = Get-ChildItem -Path $MonthFolder -Filter "*SBI*.xlsx" | Select-Object -First 1

  # -------------------------------------------------------------
  # 1. READ PHONEPE (Primary UPI Ledger)
  # -------------------------------------------------------------
  if ($ppFile) {
    Write-Output "Processing PhonePe Statement: $($ppFile.Name)"
    $tempPP = ".\_temp_pp.xlsx"
    Copy-Item $ppFile.FullName $tempPP -Force
    $wbPP = $excel.Workbooks.Open((Resolve-Path $tempPP).Path, 0, $true)
    $shPP = $wbPP.Sheets.Item(1)
    $ppRows = $shPP.UsedRange.Rows.Count

    for ($r = 4; $r -le $ppRows; $r++) {
      $dStr    = [string]$shPP.Cells.Item($r, 1).Text
      $tStr    = [string]$shPP.Cells.Item($r, 2).Text
      $details = [string]$shPP.Cells.Item($r, 3).Text
      $txnId   = [string]$shPP.Cells.Item($r, 4).Text
      $utrVal  = [string]$shPP.Cells.Item($r, 5).Value2
      $type    = [string]$shPP.Cells.Item($r, 6).Text
      $instr   = [string]$shPP.Cells.Item($r, 7).Text
      $amtVal  = [string]$shPP.Cells.Item($r, 8).Text

      if ($dStr.Trim() -eq "" -or $details.Trim() -eq "") { continue }
      if ($details -match "automatically generated|disclaimer|terms-conditions") { break }

      $amt = [double]($amtVal -replace "[^0-9.]", "")
      $cleanUtr = ($utrVal -replace "[^0-9]", "")
      if ($cleanUtr.Length -ge 6) { [void]$utrSet.Add($cleanUtr) }

      # Mode determination
      $mode = "other"
      if ($instr -match "653029") { $mode = "cc-pp-hdfc" }
      elseif ($instr -match "2387") { $mode = "hdfc-pp" }
      elseif ($instr -match "4526") { $mode = "sbi-pp" }
      elseif ($instr -match "9928") { $mode = "bank-9928" }
      else { $mode = "pp-other" }

      # Clean Purpose
      $purpose = $details -replace "^Paid to\s+", "" -replace "^Payment to\s+", "" -replace "^Transfer to\s+", "Transfer to " -replace "^Received from\s+", "Received from "

      # Domain classification via rules.json
      $domain = Get-CategoryFromRules "$purpose $details"
      if ($domain -eq "Misc" -and $details -match "^Paid to " -and $amt -le 350) {
        $domain = "Food"
      }

      $status = "Living Expense"
      if ($domain -eq "Transfers") { $status = "Transfer" }
      elseif ($domain -eq "Investment") { $status = "Investment" }
      if ($type -match "CREDIT") { $status = "Income / Credit" }

      [void]$allTxns.Add([PSCustomObject]@{
        Timestamp = "$dStr $tStr"
        Date = $dStr
        Amount = $amt
        Purpose = $purpose
        Mode = $mode
        Domain = $domain
        Source = "PhonePe"
        Status = $status
        Type = $type.ToUpper()
        RefId = $txnId
      })
    }
    $wbPP.Close($false)
    Remove-Item $tempPP -Force
    Write-Output "PhonePe parsed: $($allTxns.Count) transactions. UTRs tracked: $($utrSet.Count)"
  }

  # -------------------------------------------------------------
  # 2. READ ICICI CREDIT CARD (Direct Card Swipes)
  # -------------------------------------------------------------
  if ($icFile) {
    Write-Output "Processing ICICI Statement: $($icFile.Name)"
    $tempIC = ".\_temp_icici.xlsx"
    Copy-Item $icFile.FullName $tempIC -Force
    $wbIC = $excel.Workbooks.Open((Resolve-Path $tempIC).Path, 0, $true)
    $shIC = $wbIC.Sheets.Item(1)
    $icRows = $shIC.UsedRange.Rows.Count
    $icCount = 0

    for ($r = 3; $r -le $icRows; $r++) {
      $dStr    = [string]$shIC.Cells.Item($r, 1).Text
      $details = [string]$shIC.Cells.Item($r, 2).Text
      $amtVal  = [string]$shIC.Cells.Item($r, 4).Text
      $sign    = [string]$shIC.Cells.Item($r, 5).Text

      if ($dStr.Trim() -eq "" -or $details.Trim() -eq "") { continue }
      $amt = [double]($amtVal -replace "[^0-9.]", "")
      $type = if ($sign.Trim() -eq "CR") { "CREDIT" } else { "DEBIT" }

      if ($details -match "BBPS Payment received") {
        $domain = "Transfers"
        $status = "CC Bill Repayment"
      } else {
        $domain = Get-CategoryFromRules $details
        $status = "Living Expense"
      }

      if ($type -eq "CREDIT" -and $status -ne "CC Bill Repayment") {
        $status = "Refund / Credit"
      }

      [void]$allTxns.Add([PSCustomObject]@{
        Timestamp = $dStr
        Date = $dStr
        Amount = $amt
        Purpose = $details
        Mode = "cc-icici"
        Domain = $domain
        Source = "ICICI-CreditCard"
        Status = $status
        Type = $type
        RefId = ""
      })
      $icCount++
    }
    $wbIC.Close($false)
    Remove-Item $tempIC -Force
    Write-Output "ICICI parsed: $icCount transactions added."
  }

  # -------------------------------------------------------------
  # 3. READ HDFC BANK (ATM, Inflows, Unique Non-UPI items)
  # -------------------------------------------------------------
  if ($hdFile) {
    Write-Output "Processing HDFC Statement: $($hdFile.Name)"
    $tempHD = ".\_temp_hdfc.xls"
    Copy-Item $hdFile.FullName $tempHD -Force
    $wbHD = $excel.Workbooks.Open((Resolve-Path $tempHD).Path, 0, $true)
    $shHD = $wbHD.Sheets.Item(1)
    $hdRows = $shHD.UsedRange.Rows.Count
    $hdAdded = 0
    $hdSkipped = 0

    for ($r = 3; $r -le $hdRows; $r++) {
      $dStr = [string]$shHD.Cells.Item($r, 1).Text
      $narr = [string]$shHD.Cells.Item($r, 2).Text
      $wStr = [string]$shHD.Cells.Item($r, 4).Text
      $dStrAmt = [string]$shHD.Cells.Item($r, 5).Text

      if ($dStr.Trim() -eq "" -or $narr.Trim() -eq "" -or $narr -match "\*\*\*") { continue }

      $wAmt = if ($wStr.Trim() -ne "") { [double]($wStr -replace "[^0-9.]", "") } else { 0 }
      $dAmt = if ($dStrAmt.Trim() -ne "") { [double]($dStrAmt -replace "[^0-9.]", "") } else { 0 }
      $amt = if ($wAmt -gt 0) { $wAmt } else { $dAmt }
      $type = if ($wAmt -gt 0) { "DEBIT" } else { "CREDIT" }

      # Deduplication: check if UTR exists in PhonePe
      $isDuplicate = $false
      foreach ($u in $utrSet) {
        if ($narr.Contains($u)) { $isDuplicate = $true; break }
      }
      if ($isDuplicate) {
        $hdSkipped++
        continue
      }

      $mode = "hdfc-pp"
      $domain = "Misc"
      $status = "Living Expense"
      $purpose = $narr

      if ($narr -match "^ATW-") {
        $mode = "Cash"
        $domain = "Misc"
        $purpose = "ATM Cash Withdrawal"
        $status = "ATM Withdrawal"
      } elseif ($narr -match "CRED CLUB") {
        $mode = "hdfc-pp"
        $domain = "Transfers"
        $purpose = "Credit Card Bill (CRED)"
        $status = "CC Bill Repayment"
      } elseif ($narr -match "NATIONAL HIGHWAYS AU") {
        $mode = "hdfc-pp"
        $domain = "Travel"
        $purpose = "NHAI Annual Pass"
        $status = "Living Expense"
      } elseif ($narr -match "ABHISHEK SANJAY TELS") {
        $mode = "hdfc-pp"
        $domain = "Transfers"
        $purpose = "Transfer - Abhishek"
        $status = if ($type -eq "CREDIT") { "Income / Credit" } else { "Transfer" }
      } else {
        $domain = Get-CategoryFromRules $narr
      }

      [void]$allTxns.Add([PSCustomObject]@{
        Timestamp = $dStr
        Date = $dStr
        Amount = $amt
        Purpose = $purpose
        Mode = $mode
        Domain = $domain
        Source = "HDFC-Bank"
        Status = $status
        Type = $type
        RefId = ""
      })
      $hdAdded++
    }
    $wbHD.Close($false)
    Remove-Item $tempHD -Force
    Write-Output "HDFC deduplicated: $hdSkipped UPI transactions linked to PhonePe. $hdAdded unique bank items added."
  }

  # -------------------------------------------------------------
  # BUILD CONSOLIDATED EXCEL WORKBOOK
  # -------------------------------------------------------------
  Write-Output "Generating consolidated Excel file: $OutputFile"
  $wbNew = $excel.Workbooks.Add()

  # Sheet 1: Expenditure
  $shExp = $wbNew.Sheets.Item(1)
  $shExp.Name = "Expenditure"

  $headers = @("Timestamp", "Date", "Amount", "Purpose", "Mode", "Domain", "Source", "Status", "Type", "RefId")
  for ($c = 0; $c -lt $headers.Count; $c++) {
    $cell = $shExp.Cells.Item(1, $c + 1)
    $cell.Value = $headers[$c]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 14277081
  }

  $rowIdx = 2
  foreach ($t in $allTxns) {
    $shExp.Cells.Item($rowIdx, 1).Value = $t.Timestamp
    $shExp.Cells.Item($rowIdx, 2).Value = $t.Date
    $shExp.Cells.Item($rowIdx, 3).Value = [double]$t.Amount
    $shExp.Cells.Item($rowIdx, 4).Value = $t.Purpose
    $shExp.Cells.Item($rowIdx, 5).Value = $t.Mode
    $shExp.Cells.Item($rowIdx, 6).Value = $t.Domain
    $shExp.Cells.Item($rowIdx, 7).Value = $t.Source
    $shExp.Cells.Item($rowIdx, 8).Value = $t.Status
    $shExp.Cells.Item($rowIdx, 9).Value = $t.Type
    $shExp.Cells.Item($rowIdx, 10).Value = $t.RefId
    $rowIdx++
  }
  $shExp.Columns.AutoFit()

  # Sheet 2: Bal-Aug26 (Summary & Reconciliation)
  $shBal = $wbNew.Sheets.Add([System.Reflection.Missing]::Value, $shExp)
  $shBal.Name = "Bal-Aug26"

  # Table 1: Mode Breakdown
  $shBal.Cells.Item(1, 1).Value = "Mode Breakdown (Living Spend)"
  $shBal.Cells.Item(1, 1).Font.Bold = $true
  $shBal.Cells.Item(2, 1).Value = "Mode"
  $shBal.Cells.Item(2, 2).Value = "Transactions"
  $shBal.Cells.Item(2, 3).Value = "Total Spend (INR)"
  $shBal.Cells.Item(2, 1).Font.Bold = $true
  $shBal.Cells.Item(2, 2).Font.Bold = $true
  $shBal.Cells.Item(2, 3).Font.Bold = $true

  $modeGroups = $allTxns | Where-Object { $_.Status -eq "Living Expense" -and $_.Type -eq "DEBIT" } | Group-Object Mode
  $mRow = 3
  foreach ($mg in ($modeGroups | Sort-Object Count -Descending)) {
    $tot = ($mg.Group | Measure-Object Amount -Sum).Sum
    $shBal.Cells.Item($mRow, 1).Value = [string]$mg.Name
    $shBal.Cells.Item($mRow, 2).Value = [double]$mg.Count
    $shBal.Cells.Item($mRow, 3).Value = [double]$tot
    $mRow++
  }

  # Table 2: Domain Breakdown
  $shBal.Cells.Item(1, 5).Value = "Domain Breakdown (Living Spend)"
  $shBal.Cells.Item(1, 5).Font.Bold = $true
  $shBal.Cells.Item(2, 5).Value = "Domain"
  $shBal.Cells.Item(2, 6).Value = "Transactions"
  $shBal.Cells.Item(2, 7).Value = "Total Spend (INR)"
  $shBal.Cells.Item(2, 5).Font.Bold = $true
  $shBal.Cells.Item(2, 6).Font.Bold = $true
  $shBal.Cells.Item(2, 7).Font.Bold = $true

  $domainGroups = $allTxns | Where-Object { $_.Status -eq "Living Expense" -and $_.Type -eq "DEBIT" } | Group-Object Domain
  $dRow = 3
  foreach ($dg in ($domainGroups | Sort-Object Count -Descending)) {
    $tot = ($dg.Group | Measure-Object Amount -Sum).Sum
    $shBal.Cells.Item($dRow, 5).Value = [string]$dg.Name
    $shBal.Cells.Item($dRow, 6).Value = [double]$dg.Count
    $shBal.Cells.Item($dRow, 7).Value = [double]$tot
    $dRow++
  }

  # Table 3: Reconciliation & Cash Flow
  $shBal.Cells.Item(1, 9).Value = "Cash Flow & Reconciliation"
  $shBal.Cells.Item(1, 9).Font.Bold = $true
  $shBal.Cells.Item(2, 9).Value = "Category"
  $shBal.Cells.Item(2, 10).Value = "Total Amount (INR)"
  $shBal.Cells.Item(2, 9).Font.Bold = $true
  $shBal.Cells.Item(2, 10).Font.Bold = $true

  $atmTot = ($allTxns | Where-Object { $_.Status -eq "ATM Withdrawal" } | Measure-Object Amount -Sum).Sum
  $ccTot  = ($allTxns | Where-Object { $_.Status -eq "CC Bill Repayment" } | Measure-Object Amount -Sum).Sum
  $invTot = ($allTxns | Where-Object { $_.Status -eq "Investment" } | Measure-Object Amount -Sum).Sum
  $liveTot = ($allTxns | Where-Object { $_.Status -eq "Living Expense" -and $_.Type -eq "DEBIT" } | Measure-Object Amount -Sum).Sum
  $inflowTot = ($allTxns | Where-Object { $_.Status -eq "Income / Credit" } | Measure-Object Amount -Sum).Sum

  $reconcil = @(
    @("Living Expenses (Actual Spend)", [double]$liveTot),
    @("ATM Cash Withdrawn", [double]$atmTot),
    @("Credit Card Bill Payments", [double]$ccTot),
    @("Investments", [double]$invTot),
    @("Inflows / Deposits", [double]$inflowTot)
  )

  $rRow = 3
  foreach ($item in $reconcil) {
    $shBal.Cells.Item($rRow, 9).Value = [string]$item[0]
    $shBal.Cells.Item($rRow, 10).Value = [double]$item[1]
    $rRow++
  }

  $shBal.Columns.AutoFit()

  if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
  $wbNew.SaveAs($OutputFile, 51)
  $wbNew.Close($false)

  Write-Output "Consolidation complete!"
  Write-Output "Saved to: $OutputFile"
  Write-Output "Total Living Expenses: INR $liveTot"
} catch {
  Write-Output "Error during consolidation: $_"
  Write-Output "At line: $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line)"
  Write-Output "$($_.ScriptStackTrace)"
} finally {
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
