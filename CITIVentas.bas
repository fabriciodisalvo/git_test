Attribute VB_Name = "CITIVentas"
'******************************************************************************
' Module        : User Macro
' Description   : This module sheet validates and creates a unique table with
'               : all purchases from the input table.
' Include       : None
' Interface     : InputValidationRules
'               : CreateUniqueVATTable
' Comments      :
'               :
'******************************************************************************

Dim sourcesheet As Worksheet, destsheet As Worksheet
Dim LastRow As Long, LastColumn As Long

Private Sub CreateUniqueVATTable()
'******************************************************************************
' Description   : This sub creates an unique table for monthly purchases.
' Input         : inputSheet - Raw data from SAP
' Output        : exportSheet - Summarize table.
' Comments      :
' History       :
'   #.001   Column by column, manual calculation for each concept.
'   #.021   Simplifyied creation of table.
'   #.000
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
        For I = 2 To 5:
            Condition = Sheets("reference").Range("F" & I)
            rolling_sum = rolling_sum + WorksheetFunction.SumIfs(sourcesheet.Range("I2:I" & LastRow), _
                sourcesheet.Range("E2:E" & LastRow), destsheet.Range("C" & n), _
                sourcesheet.Range("F2:F" & LastRow), destsheet.Range("D" & n), _
                sourcesheet.Range("H2:H" & LastRow), Condition _
                )
        Next I
        destsheet.Range("G" & n) = rolling_sum * (-1)
    ' Populate EXENTO amount in table
        rolling_sum = 0
        For k = 2 To 2:
            Condition = Sheets("reference").Range("H" & k)
            rolling_sum = rolling_sum + WorksheetFunction.SumIfs(sourcesheet.Range("I2:I" & LastRow), _
                sourcesheet.Range("E2:E" & LastRow), destsheet.Range("C" & n), _
                sourcesheet.Range("F2:F" & LastRow), destsheet.Range("D" & n), _
                sourcesheet.Range("H2:H" & LastRow), Condition _
                )
        Next k
        destsheet.Range("H" & n) = rolling_sum * (-1)
    ' Populate VAT amounts in table (7 Columns, 4 to L)
        For k = 1 To 4:
            Condition = destsheet.Cells(7, k + 8)
            destsheet.Cells(n, k + 8) = WorksheetFunction.SumIfs(sourcesheet.Range("J2:J" & LastRow), _
                sourcesheet.Range("E2:E" & LastRow), destsheet.Range("C" & n), _
                sourcesheet.Range("F2:F" & LastRow), destsheet.Range("D" & n), _
                sourcesheet.Range("H2:H" & LastRow), Condition _
                ) * (-1)
        Next k
    ' Populate TAX amounts in table (3 Columns, M to O)
        For k = 1 To 3:
            Condition = destsheet.Cells(7, k + 12)
            destsheet.Cells(n, k + 12) = WorksheetFunction.SumIfs(sourcesheet.Range("I2:I" & LastRow), _
                sourcesheet.Range("E2:E" & LastRow), destsheet.Range("C" & n), _
                sourcesheet.Range("F2:F" & LastRow), destsheet.Range("D" & n), _
                sourcesheet.Range("H2:H" & LastRow), Condition _
                ) * (-1)
        Next k
    ' Populate TOTAL amount in table
        destsheet.Range("P" & n).FormulaR1C1 = "=SUM(RC[-9]:RC[-1])"

    Next n

End Sub

Private Sub CreateFLATFILES()

'******************************************************************************
' Description   : This sub creates the CITI flat files.
' Input         : exportSheet - Summarize table.
' Output        : REGINFO_CV_COMPRAS_CBTE
' Comments      :
' History       :
'   #.001   Column by column, manual calculation for each concept.
'   #.021   Tree-like calculation of fields.
'   #.030   All Outputs completed
'******************************************************************************
Dim LastSummarizeTableRow As Long

LastSummarizeTableRow = destsheet.Cells(destsheet.Rows.Count, "B").End(xlUp).Row

base_index = 0
vat_index = 0
import_index = 0
Sheets("CITI_VENTAS_CBTE").Range("A:A").ClearContents
Sheets("CITI_VENTAS_ALICUOTAS").Range("A:A").ClearContents


For n = 8 To LastSummarizeTableRow:

'   Create blank fields for each necessary line. Get info on non-variable fields, like DATE and DOCUMENT TYPE, for BASE flat file.
    complete_string = ""
    fecha_de_comprobante = Format(Range("E" & n), "YYYYMMDD")               ' From Table
    tipo_de_comprobante = Application.VLookup(Left(Range("D" & n), 3), Sheets("reference").Range("$A$2:$E$1003"), 2, False)
    punto_de_venta = "0"                                                    ' Variable, depending on tipo_de_comprobante
    numero_de_comprobante = "0"                                             ' Variable, depending on tipo_de_comprobante
    numero_de_comprobante_hasta = "0"                                       ' Variable, depending on tipo_de_comprobante
    codigo_de_documento_del_comprador = "80"                                ' FIXED
    numero_de_documento_del_comprador = Range("C" & n)                      ' From Table
    denominacion_del_comprador = Range("F" & n)                             ' From Table
    importe_total_de_la_operacion = Format(Abs(Round(Range("P" & n), 2) * 100), String(15, "0"))
    importe_no_neto_gravado = Format("00", String(15, "0"))                 ' Calculation
    percepcion_no_categorizados = "0"                                       ' Calculation
    importe_exento = Format(Abs(Round(Range("H" & n), 2) * 100), String(15, "0")) ' Calculation
    percepcion_impuestos_nacionales = "0"                                   ' Calculation
    percepcion_iibb = "0"                                                   ' Calculation
    percepcion_impuestos_municipales = "0"                                  ' Calculation
    percepcion_impuestos_internos = "0"                                     ' Calculation
    codigo_de_moneda = "PES"                                                ' FIXED
    tipo_de_cambio = "1000000"                                              ' FIXED
    cantidad_de_alicuotas_de_iva = 0                                        ' Calculation
    codigo_de_operacion = "0"                                               ' Calculation
    credito_fiscal_computable = 0                                           ' Calculation
    otros_tributos = "0"                                                    ' Calculation
    fecha_vencimiento_pago = "0"                                            ' NOT USED
    
'   Create blank fields for each necessary line for VAT flat file.
    vat_complete_string = ""
    importe_neto_gravado = 0                                                ' Calculation
    alicuota_de_iva = "0"                                                   ' From Table
    impuesto_liquidado = 0                                                  ' From Table
    
'   Get DOCUMENT NUMBERING for FACTURAS, NOTAS DE DEBITO, NOTAS DE CREDITO.
    If Not IsError(Application.Match(tipo_de_comprobante, Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 19, 20, 21), False)) Then
        punto_de_venta = Mid(Range("D" & n), 4, 4)
        numero_de_comprobante = Mid(Range("D" & n), 9, 8)
        numero_de_comprobante_hasta = numero_de_comprobante
    End If

'   Get VAT information for creating lines for this and additional flat files.
'   Check 21% VAT information.
    vat_value = Range("I" & n).Value
    If vat_value <> 0 Then
        cantidad_de_alicuotas_de_iva = cantidad_de_alicuotas_de_iva + 1
        credito_fiscal_computable = credito_fiscal_computable + Abs(Round(vat_value, 2) * 100)
        importe_neto_gravado = Abs(Round(vat_value, 2) / 0.21 * 100)
        alicuota_de_iva = "0005"
        impuesto_liquidado = Abs(Round(vat_value, 2) * 100)
        vat_complete_string = Format(tipo_de_comprobante, String(3, "0")) & _
                                    Format(punto_de_venta, String(5, "0")) & _
                                    Format(numero_de_comprobante, String(20, "0")) & _
                                    Format(codigo_de_documento_del_vendedor, String(2, "0")) & _
                                    Format(numero_de_documento_del_vendedor, String(20, "0")) & _
                                    Format(importe_neto_gravado, String(15, "0")) & _
                                    Format(alicuota_de_iva, String(4, "0")) & _
                                    Format(impuesto_liquidado, String(15, "0"))
        vat_index = vat_index + 1
        Sheets("CITI_VENTAS_ALICUOTAS").Range("A" & vat_index).Value = "'" & vat_complete_string
    End If
'   Check 10.5% VAT information.
    vat_value = Range("J" & n).Value
    If vat_value <> 0 Then
        cantidad_de_alicuotas_de_iva = cantidad_de_alicuotas_de_iva + 1
        credito_fiscal_computable = credito_fiscal_computable + Abs(Round(vat_value, 2) * 100)
        importe_neto_gravado = Abs(Round(vat_value, 2) / 0.105 * 100)
        alicuota_de_iva = "0004"
        impuesto_liquidado = Abs(Round(vat_value, 2) * 100)
        vat_complete_string = Format(tipo_de_comprobante, String(3, "0")) & _
                                Format(punto_de_venta, String(5, "0")) & _
                                Format(numero_de_comprobante, String(20, "0")) & _
                                Format(codigo_de_documento_del_vendedor, String(2, "0")) & _
                                Format(numero_de_documento_del_vendedor, String(20, "0")) & _
                                Format(importe_neto_gravado, String(15, "0")) & _
                                Format(alicuota_de_iva, String(4, "0")) & _
                                Format(impuesto_liquidado, String(15, "0"))
        vat_index = vat_index + 1
        Sheets("CITI_VENTAS_ALICUOTAS").Range("A" & vat_index).Value = "'" & vat_complete_string
    End If

'   Check  0.0% VAT information (ALWAYS THE LAST TO CHECK), also update CODIGO DE OPERACION.
    If (Not IsError(Application.Match(tipo_de_comprobante, Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 19, 20, 21), False))) And cantidad_de_alicuotas_de_iva = 0 Then
        cantidad_de_alicuotas_de_iva = cantidad_de_alicuotas_de_iva + 1
        importe_neto_gravado = 0
        alicuota_de_iva = "0003"
        impuesto_liquidado = 0
        vat_complete_string = Format(tipo_de_comprobante, String(3, "0")) & _
                                Format(punto_de_venta, String(5, "0")) & _
                                Format(numero_de_comprobante, String(20, "0")) & _
                                Format(codigo_de_documento_del_vendedor, String(2, "0")) & _
                                Format(numero_de_documento_del_vendedor, String(20, "0")) & _
                                Format(importe_neto_gravado, String(15, "0")) & _
                                Format(alicuota_de_iva, String(4, "0")) & _
                                Format(impuesto_liquidado, String(15, "0"))
        vat_index = vat_index + 1
        Sheets("CITI_VENTAS_ALICUOTAS").Range("A" & vat_index).Value = "'" & vat_complete_string
        
        If (Left(Range("D" & n), 3)) = "FTZ" Or (Left(Range("D" & n), 3)) = "FCE" Then
            codigo_de_operacion = "X"
        ElseIf importe_exento > 0 Then
            codigo_de_operacion = "E"
        Else
            codigo_de_operacion = "A"
        End If
    End If
      
'   After reviewing taxes, get CODIGO DE OPERACION based on the document.
    If cantidad_de_alicuotas_de_iva = 0 Then
        If (Left(Range("D" & n), 3)) = "FTZ" Or (Left(Range("D" & n), 3)) = "FCE" Then
            codigo_de_operacion = "X"
        ElseIf importe_exento > 0 Then
            codigo_de_operacion = "E"
        Else
            codigo_de_operacion = "A"
        End If
    End If

  
    complete_string = Format(fecha_de_comprobante, String(8, "0")) & _
                        Format(tipo_de_comprobante, String(3, "0")) & _
                        Format(punto_de_venta, String(5, "0")) & _
                        Format(numero_de_comprobante, String(20, "0")) & _
                        Format(numero_de_comprobante_hasta, String(20, "0")) & _
                        Format(codigo_de_documento_del_comprador, String(2, "0")) & _
                        Format(numero_de_documento_del_comprador, String(20, "0")) & _
                        Left(denominacion_del_comprador & Space(30), 30) & _
                        importe_total_de_la_operacion & _
                        importe_no_neto_gravado & _
                        Format(percepcion_no_categorizados, String(15, "0")) & _
                        importe_exento & _
                        Format(percepcion_impuestos_nacionales, String(15, "0")) & _
                        Format(percepcion_iibb, String(15, "0")) & _
                        Format(percepcion_impuestos_municipales, String(15, "0")) & _
                        Format(percepcion_impuestos_internos, String(15, "0")) & _
                        Left(codigo_de_moneda & Space(3), 3) & _
                        Format(tipo_de_cambio, String(10, "0")) & _
                        Format(cantidad_de_alicuotas_de_iva, String(1, "0")) & _
                        Left(codigo_de_operacion & Space(1), 1) & _
                        Format(otros_tributos, String(15, "0")) & _
                        Format(fecha_vencimiento_pago, String(8, "0"))
    
    Sheets("CITI_VENTAS_CBTE").Range("A" & n - 7).Value = "'" & complete_string

Next n

'   Create actual flat files.
    Dim rRange As Range
    Dim ws As Worksheet
    Dim stTextName As String
    pop_up_message = ""

'   Create BCTE flat file.
    On Error Resume Next
    Application.DisplayAlerts = False
    Set rRange = Sheets("CITI_VENTAS_CBTE").Range("A1:A" & LastSummarizeTableRow - 7)
    On Error GoTo 0
    Application.DisplayAlerts = True

    If rRange Is Nothing Then
        pop_up_message = pop_up_message & "No hay datos para el archivo de comprobantes." & vbCrLf
    Else
        stTextName = "CITI_VENTAS_CBTE"
        stPath = ActiveWorkbook.Path
        Set ws = Worksheets.Add()
        rRange.Copy ws.Cells(1, 1)
        ws.Move
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs _
        Filename:=stPath & "\" & stTextName, _
        FileFormat:=xlText
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
        pop_up_message = pop_up_message & "Archivo de comprobantes generado." & vbCrLf
    End If

'   Create ALICUOTAS flat file.
    On Error Resume Next
    Application.DisplayAlerts = False
    Set rRange = Sheets("CITI_VENTAS_ALICUOTAS").Range("A1:A" & vat_index)
    On Error GoTo 0
    Application.DisplayAlerts = True
     
    If rRange Is Nothing Then
        pop_up_message = pop_up_message & "No hay datos en el archivo de alicuotas." & vbCrLf
    Else
        stTextName = "CITI_VENTAS_ALICUOTAS"
        stPath = ActiveWorkbook.Path
        Set ws = Worksheets.Add()
        rRange.Copy ws.Cells(1, 1)
        ws.Move
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs _
        Filename:=stPath & "\" & stTextName, _
        FileFormat:=xlText
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
        pop_up_message = pop_up_message & "Archivo de alicuotas generado." & vbCrLf
    End If

    MsgBox (pop_up_message)

End Sub


Private Sub CalculateRunTime_Seconds()
'******************************************************************************
' Description   : This sub determines seconds to run the included subs.
' Input         : sheet
' Output        : sheet
' Comments      :
' History       :
'   #.001   Source:  www.TheSpreadsheetGuru.com/the-code-vault.
'   #.002   Modified to accommodate all macros.
'   #.003
'******************************************************************************

    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    'Remember time when macro starts
      StartTime = Timer
    
    '*****************************
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    ActiveSheet.DisplayPageBreaks = False
        ' Call InputValidate()
        Sheets("export").Activate
        Call CreateUniqueVATTable
        Call CreateFLATFILES
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    ActiveSheet.DisplayPageBreaks = True
    '*****************************
    
    'Determine how many seconds code took to run
      SecondsElapsed = Round(Timer - StartTime, 2)
    
    'Notify user in seconds
    ' MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
    destsheet.Range("M6").Value = LastRow - 1 & " rows"
    destsheet.Range("N6").Value = "This code ran successfully in " & SecondsElapsed & " seconds"

End Sub

Sub RunCITIVENTAS()

    Call CalculateRunTime_Seconds

End Sub
