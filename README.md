# ValidateAndAddZeros VBA Macro

## Overview

`ValidateAndAddZeros` is a VBA macro designed for Excel users managing product data. It validates European Article Number (EAN) codes in column A, ensures each valid code is padded to 13 digits, and **clears any invalid entries** based on custom business rules.

This macro helps maintain a clean dataset of product identifiers by enforcing standardized EAN formatting.

---

## Features

- 🔍 Scans column A (starting from row 2) for EAN values.
- ✅ Validates EANs using specific business rules:
  - Must be numeric.
  - Must not exceed 13 digits.
  - Must **not** start with "2".
  - Must **not** include "000" or "00000" in critical positions.
- 🔢 Automatically adds leading zeros to make valid EANs exactly 13 digits.
- ❌ **Clears** cells containing invalid or corrupt EANs.

---

## Usage

1. Open your Excel file containing product EANs.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module or paste this macro into `Sheet1`.
4. Run the `ValidateAndAddZeros` macro.
5. It will process data in **column A**, starting from row 2.

---

## Code

```vba
Sub ValidateAndAddZeros()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim eanRange As Range
    Dim cell As Range

    ' Set the active sheet
    Set ws = ActiveSheet

    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Set the range of EAN codes
    Set eanRange = ws.Range("A2:A" & lastRow)

    ' Loop through each cell in the range
    For Each cell In eanRange
        Dim ean As String
        ean = CStr(cell.Value)

        ' Check validity of EANs and add zeros
        If IsNumeric(ean) And Len(ean) <= 13 And Left(ean, 1) <> "2" And _
           Mid(ean, 1, 3) <> "000" And Mid(ean, 3, 5) <> "00000" And Mid(ean, 8, 5) <> "00000" Then

            ' Add leading zeros to make it 13 digits
            Dim newEAN As String
            newEAN = WorksheetFunction.Rept("0", 13 - Len(ean)) & ean

            ' Validate new EAN
            If IsNumeric(newEAN) And Len(newEAN) = 13 And _
               Left(newEAN, 1) <> "2" And Mid(newEAN, 1, 3) <> "000" And _
               Mid(newEAN, 3, 5) <> "00000" And Mid(newEAN, 8, 5) <> "00000" Then
                cell.Value = newEAN
            Else
                cell.ClearContents
            End If
        Else
            cell.ClearContents
        End If
    Next cell
End Sub
