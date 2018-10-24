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
'   #2   Simplifyied creation of table.
'   #3
'******************************************************************************

    Set sourcesheet = Sheets("input")
    Set destsheet = Sheets("export")
    
    ' Create unique list of documents
    LastRow = sourcesheet.Cells(sourcesheet.Rows.Count, "A").End(xlUp).Row
    destsheet.Range("B8:P999").ClearContents
    sourcesheet.Range("E2:F" & LastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=destsheet.Range("C8:D999"), Unique:=True
    SummLastRow = destsheet.Cells(sourcesheet.Rows.Count, "C").End(xlUp).Row
    
    For n = 8 To SummLastRow:
    ' Populate INDEX in table
        destsheet.Range("B" & n) = n - 7
    ' Populate DATE in table
        destsheet.Range("E" & n) = WorksheetFunction.MinIfs(sourcesheet.Range("A2:A" & LastRow), _
            sourcesheet.Range("E2:E" & LastRow), destsheet.Range("C" & n), _
            sourcesheet.Range("F2:F" & LastRow), destsheet.Range("D" & n) _
            )
    ' Populate VENDOR in table
        destsheet.Range("F" & n) = Application.Index(sourcesheet.Range("D2:D999"), Application.Match(destsheet.Range("C" & n), sourcesheet.Range("E2:E999"), 0), 1)
    ' Populate NETO amount in table
        rolling_sum = 0
        For i = 91 To 94:
            Condition = Sheets("data").Range("D" & i)
            rolling_sum = rolling_sum + WorksheetFunction.SumIfs(sourcesheet.Range("I2:I" & LastRow), _
                sourcesheet.Range("E2:E" & LastRow), destsheet.Range("C" & n), _
                sourcesheet.Range("F2:F" & LastRow), destsheet.Range("D" & n), _
                sourcesheet.Range("H2:H" & LastRow), Condition _
                )
        Next i
        destsheet.Range("G" & n) = rolling_sum
    ' Populate EXENTO amount in table
        rolling_sum = 0
        For k = 97 To 97:
            Condition = Sheets("data").Range("D" & k)
            rolling_sum = rolling_sum + WorksheetFunction.SumIfs(sourcesheet.Range("I2:I" & LastRow), _
                sourcesheet.Range("E2:E" & LastRow), destsheet.Range("C" & n), _
                sourcesheet.Range("F2:F" & LastRow), destsheet.Range("D" & n), _
                sourcesheet.Range("H2:H" & LastRow), Condition _
                )
        Next k
        destsheet.Range("H" & n) = rolling_sum
    ' Populate VAT amounts in table (7 Columns, 4 to L)
        For k = 1 To 4:
            Condition = destsheet.Cells(7, k + 8)
            destsheet.Cells(n, k + 8) = WorksheetFunction.SumIfs(sourcesheet.Range("J2:J" & LastRow), _
                sourcesheet.Range("E2:E" & LastRow), destsheet.Range("C" & n), _
                sourcesheet.Range("F2:F" & LastRow), destsheet.Range("D" & n), _
                sourcesheet.Range("H2:H" & LastRow), Condition _
                )
        Next k
    ' Populate TAX amounts in table (3 Columns, M to O)
        For k = 1 To 3:
            Condition = destsheet.Cells(7, k + 12)
            destsheet.Cells(n, k + 12) = WorksheetFunction.SumIfs(sourcesheet.Range("I2:I" & LastRow), _
                sourcesheet.Range("E2:E" & LastRow), destsheet.Range("C" & n), _
                sourcesheet.Range("F2:F" & LastRow), destsheet.Range("D" & n), _
                sourcesheet.Range("H2:H" & LastRow), Condition _
                )
        Next k
    ' Populate TOTAL amount in table
        destsheet.Range("P" & n).FormulaR1C1 = "=SUM(RC[-9]:RC[-1])"

    Next n

End Sub

Sub CreateCPBT()
'******************************************************************************
' Description   : This sub creates the CBTE flat file.
' Input         : exportSheet - Summarize table.
' Output        : REGINFO_CV_COMPRAS_CBTE
' Comments      :
' History       :
'   #1   Column by column, manual calculation for each concept.
'   #2
'   #3
'******************************************************************************
For n = 8 To 181:
    complete_string = ""
    
    fecha_de_comprobante = Format(Range("E" & n), "YYYYMMDD")
    tipo_de_comprobante = WorksheetFunction.VLookup(Left(Range("D" & n), 3), Sheets("data").Range("$B$2:$D$1003"), 2, False)
    
    If tipo_de_comprobante < 66 Then
        punto_de_venta = Mid(Range("D" & n), 4, 4)
        numero_de_comprobante = Mid(Range("D" & n), 9, 99)
        despacho = ""
    ElseIf tipo_de_comprobante = 66 Then
        punto_de_venta = 0
        numero_de_comprobante = 0
        despacho = Range("D" & n)
    ElseIf tipo_de_comprobante = 99 Then
        punto_de_venta = 0
        numero_de_comprobante = Mid(Range("D" & n), 4, 99)
        despacho = ""
    Else
        punto_de_venta = 0
        numero_de_comprobante = Mid(Range("D" & n), 9, 99)
        despacho = ""
    End If
    
    codigo_de_documento_del_vendedor = "80"
    numero_de_documento_del_vendedor = Range("C" & n)
    denominacion_del_vendedor = Range("F" & n)
    importe_total_de_la_operacion = Abs(Round(Range("P" & n), 2) * 100)
    
    If Range("H" & n) < 0 Then
        importe_no_neto_gravado = Abs(Round(Range("H" & n), 2) * 100)
    Else
        importe_no_neto_gravado = "0"
    End If
    If (tipo_de_comprobante > 5 And tipo_de_comprobante < 17) Or tipo_de_comprobante = 66 Then
        If Mid(despacho, 6, 1) = "E" Then
            importe_exento = Abs(Round(Range("H" & n), 2) * 100)
        Else
            importe_exento = "0"
        End If
    Else
        importe_exento = Abs(Round(Range("H" & n), 2) * 100)
    End If
    
    importe_percepciones_iva = Abs(Round(Range("O" & n), 2) * 100)
    importe_percepciones_otros = Abs(Round(Range("N" & n), 2) * 100)
    importe_percepciones_iibb = Abs(Round(Range("M" & n), 2) * 100)
    importe_percepciones_municipales = "0"
    importe_impuestos_internos = "0"
    codigo_de_moneda = "PES"
    tipo_de_cambio = "1000000"
    
    alicuotas_count = 0
    If Range("I" & n) <> 0 Then alicuotas_count = alicoutas_count + 1
    If Range("J" & n) <> 0 Then alicuotas_count = alicoutas_count + 1
    If Range("K" & n) <> 0 Then alicuotas_count = alicoutas_count + 1
    If Range("L" & n) <> 0 Then alicuotas_count = alicoutas_count + 1
    cantidad_de_alicuotas_de_iva = alicuotas_count
    
    codigo_de_operacion = "CHECK"
    credito_fiscal_computable = Abs(Round(Range("I" & n) + Range("J" & n) + Range("K" & n) + Range("L" & n), 2) * 100)
    otros_tributos = "0"
    CUIT_emisor_corredor = "0"
    denominacion_emisor_corredor = ""
    iva_comision = "0"
    
    complete_string = fecha_de_comprobante & _
                        Format(tipo_de_comprobante, String(3, "0")) & _
                        Format(punto_de_venta, String(5, "0")) & _
                        Format(numero_de_comprobante, String(20, "0")) & _
                        Left(despacho & Space(16), 16) & _
                        Format(codigo_de_documento_del_vendedor, String(2, "0")) & _
                        Format(numero_de_documento_del_vendedor, String(20, "0")) & _
                        Left(denominacion_del_vendedor & Space(30), 30) & _
                        Format(importe_total_de_la_operacion, String(15, "0")) & _
                        Format(importe_no_neto_gravado, String(15, "0")) & _
                        Format(importe_exento, String(15, "0")) & _
                        Format(importe_percepciones_iva, String(15, "0")) & _
                        Format(importe_percepciones_otros, String(15, "0")) & _
                        Format(importe_percepciones_iibb, String(15, "0")) & _
                        Format(importe_percepciones_municipales, String(15, "0")) & _
                        Format(importe_impuestos_internos, String(15, "0")) & _
                        Left(codigo_de_moneda & Space(3), 3) & _
                        Format(tipo_de_cambio, String(10, "0")) & _
                        Format(cantidad_de_alicuotas_de_iva, String(1, "0")) & _
                        Left(codigo_de_operacion & Space(1), 1) & _
                        Format(credito_fiscal_computable, String(15, "0")) & _
                        Format(otros_tributos, String(15, "0")) & _
                        Format(CUIT_emisor_corredor, String(11, "0")) & _
                        Left(denominacion_emisor_corredor & Space(30), 30) & _
                        Format(iva_comision, String(15, "0"))
    
    Sheets("CBTE").Range("A" & n - 7).Value = complete_string
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
Call CreateUniqueVATTable
'*****************************

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
'MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

destsheet.Range("B2").Value = LastRow - 1 & " rows"
destsheet.Range("C2").Value = "This code ran successfully in " & SecondsElapsed & " seconds"

End Sub
