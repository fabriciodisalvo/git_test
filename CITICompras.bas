Attribute VB_Name = "CITICompras"
'******************************************************************************
' Module        : User Macro
' Description   : This module sheet validates and creates a unique table with
'               : all purchases from the input table.
' Include       : None
' Interface     : InputValidationRules
'               : CreateUniqueVATTable
' Comments      :
'               :
' $History: $
'******************************************************************************


Dim sourcesheet As Worksheet, destsheet As Worksheet
Dim LastRow As Long, LastColumn As Long


Sub InputValidationRules()

End Sub


Sub CreateUniqueVATTable()
'******************************************************************************
' Description   : This sub creates an unique table for monthly purchases.
' Input         : inputSheet - Raw data from SAP
' Output        : exportSheet - Summarize table.
' Comments      :
' History       :
'   #1   Column by column, manual calculation for each concept.
'   #2
'   #3
'******************************************************************************
Set sourcesheet = Sheets("input")
Set destsheet = Sheets("export")
LastRow = sourcesheet.Cells(sourcesheet.Rows.Count, "A").End(xlUp).Row
destsheet.Range("S8:T999").ClearContents
sourcesheet.Range("E2:F" & LastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=destsheet.Range("S8:T999"), Unique:=True

For n = 8 To 200:

' Populate DATE in table
    destsheet.Range("U" & n) = WorksheetFunction.MinIfs(sourcesheet.Range("A2:A" & LastRow), _
        sourcesheet.Range("E2:E" & LastRow), destsheet.Range("S" & n), _
        sourcesheet.Range("F2:F" & LastRow), destsheet.Range("T" & n) _
        )
' Populate VENDOR in table
    destsheet.Range("V" & n) = Application.Index(sourcesheet.Range("D2:D999"), Application.Match(destsheet.Range("S" & n), sourcesheet.Range("E2:E999"), 0), 1)

' Populate NETO amount in table
    rolling_sum = 0
    For i = 91 To 94:
        Condition = Sheets("data").Range("D" & i)
        rolling_sum = rolling_sum + WorksheetFunction.SumIfs(sourcesheet.Range("I2:I" & LastRow), _
            sourcesheet.Range("E2:E" & LastRow), destsheet.Range("S" & n), _
            sourcesheet.Range("F2:F" & LastRow), destsheet.Range("T" & n), _
            sourcesheet.Range("H2:H" & LastRow), Condition _
            )
    Next i
    destsheet.Range("W" & n) = rolling_sum

' Populate EXENTO amount in table
    rolling_sum = 0
    For i = 97 To 97:
        Condition = Sheets("data").Range("D" & i)
        rolling_sum = rolling_sum + WorksheetFunction.SumIfs(sourcesheet.Range("I2:I" & LastRow), _
            sourcesheet.Range("E2:E" & LastRow), destsheet.Range("S" & n), _
            sourcesheet.Range("F2:F" & LastRow), destsheet.Range("T" & n), _
            sourcesheet.Range("H2:H" & LastRow), Condition _
            )
    Next i
    destsheet.Range("X" & n) = rolling_sum

' Populate VAT21 amount in table
    Condition = destsheet.Range("Y7")
    destsheet.Range("Y" & n) = WorksheetFunction.SumIfs(sourcesheet.Range("J2:J" & LastRow), _
        sourcesheet.Range("E2:E" & LastRow), destsheet.Range("S" & n), _
        sourcesheet.Range("F2:F" & LastRow), destsheet.Range("T" & n), _
        sourcesheet.Range("H2:H" & LastRow), Condition _
        )
        
' Populate VAT105 amount in table
    Condition = destsheet.Range("Z7")
    destsheet.Range("Z" & n) = WorksheetFunction.SumIfs(sourcesheet.Range("J2:J" & LastRow), _
        sourcesheet.Range("E2:E" & LastRow), destsheet.Range("S" & n), _
        sourcesheet.Range("F2:F" & LastRow), destsheet.Range("T" & n), _
        sourcesheet.Range("H2:H" & LastRow), Condition _
        )
                
' Populate VAT27 amount in table
    Condition = destsheet.Range("AA7")
    destsheet.Range("AA" & n) = WorksheetFunction.SumIfs(sourcesheet.Range("J2:J" & LastRow), _
        sourcesheet.Range("E2:E" & LastRow), destsheet.Range("S" & n), _
        sourcesheet.Range("F2:F" & LastRow), destsheet.Range("T" & n), _
        sourcesheet.Range("H2:H" & LastRow), Condition _
        )
                
' Populate VAT10A amount in table
    Condition = destsheet.Range("AB7")
    destsheet.Range("AB" & n) = WorksheetFunction.SumIfs(sourcesheet.Range("J2:J" & LastRow), _
        sourcesheet.Range("E2:E" & LastRow), destsheet.Range("S" & n), _
        sourcesheet.Range("F2:F" & LastRow), destsheet.Range("T" & n), _
        sourcesheet.Range("H2:H" & LastRow), Condition _
        )
                
' Populate IIBB amount in table
    Condition = destsheet.Range("AC7")
    destsheet.Range("AC" & n) = WorksheetFunction.SumIfs(sourcesheet.Range("J2:J" & LastRow), _
        sourcesheet.Range("E2:E" & LastRow), destsheet.Range("S" & n), _
        sourcesheet.Range("F2:F" & LastRow), destsheet.Range("T" & n), _
        sourcesheet.Range("H2:H" & LastRow), Condition _
        )
                
' Populate INCOME TAX amount in table
    Condition = destsheet.Range("AD7")
    destsheet.Range("AD" & n) = WorksheetFunction.SumIfs(sourcesheet.Range("J2:J" & LastRow), _
        sourcesheet.Range("E2:E" & LastRow), destsheet.Range("S" & n), _
        sourcesheet.Range("F2:F" & LastRow), destsheet.Range("T" & n), _
        sourcesheet.Range("H2:H" & LastRow), Condition _
        )
                
' Populate VAT TAX amount in table
    Condition = destsheet.Range("AE7")
    destsheet.Range("AE" & n) = WorksheetFunction.SumIfs(sourcesheet.Range("J2:J" & LastRow), _
        sourcesheet.Range("E2:E" & LastRow), destsheet.Range("S" & n), _
        sourcesheet.Range("F2:F" & LastRow), destsheet.Range("T" & n), _
        sourcesheet.Range("H2:H" & LastRow), Condition _
        )

' Populate TOTAL amount in table
    destsheet.Range("AF" & n).FormulaR1C1 = "=SUM(RC[-9]:RC[-1])"

Next n

End Sub




Sub CalculateRunTime_Seconds()
'PURPOSE: Determine how many seconds it took for code to completely run
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
  StartTime = Timer

'*****************************
Call externalLookupHOR
'*****************************

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
'MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

Range("D2").Value = LastRow - 1 & " rows"
Range("E2").Value = "This code ran successfully in " & SecondsElapsed & " seconds"

End Sub
