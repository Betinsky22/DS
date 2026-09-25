Option Explicit

Public Sub Porcentaje_retorno_Ambos()

    Dim wb As Workbook
    Dim wsRecti As Worksheet, wsPartidas As Worksheet
    Dim wsCasos As Worksheet, wsAF As Worksheet, wsEnc As Worksheet
    Dim ultimaRecti As Long, ultimaPartidas As Long
    Dim ultimaAF As Long, ultimaEnc As Long
    Dim i As Long
    Dim calculoAnterior As XlCalculation

    On Error GoTo ManejarError

    'ActiveWorkbook permite ejecutar la macro desde PERSONAL.XLSB
    'sobre el archivo de Data Stage que esté abierto y activo.
    Set wb = ActiveWorkbook

    If wb Is Nothing Then Err.Raise vbObjectError + 100, , _
        "No hay un libro activo."

    If Not ExisteHoja(wb, "Recti") _
       Or Not ExisteHoja(wb, "Partidas") _
       Or Not ExisteHoja(wb, "Casos_Ped") _
       Or Not ExisteHoja(wb, "Enc") Then
        Err.Raise vbObjectError + 101, , _
            "Falta alguna hoja de origen: Recti, Partidas, Casos_Ped o Enc."
    End If

    If ExisteHoja(wb, "Casos AF") _
       Or ExisteHoja(wb, "Importaciones MP") _
       Or ExisteHoja(wb, "Exportaciones MP") Then
        Err.Raise vbObjectError + 102, , _
            "Ya existe una hoja de resultado. Ejecuta la macro sobre una copia limpia del Data Stage."
    End If

    Set wsRecti = wb.Worksheets("Recti")
    Set wsPartidas = wb.Worksheets("Partidas")
    Set wsCasos = wb.Worksheets("Casos_Ped")
    Set wsEnc = wb.Worksheets("Enc")

    If LCase$(Trim$(CStr(wsRecti.Range("A1").Value))) = "original" _
       Or LCase$(Trim$(CStr(wsPartidas.Range("A1").Value))) = "llave" _
       Or LCase$(Trim$(CStr(wsEnc.Range("A1").Value))) = "llave" Then
        Err.Raise vbObjectError + 103, , _
            "Alguna hoja de origen ya fue modificada. Usa una copia limpia."
    End If

    calculoAnterior = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '1. RECTI: llaves del pedimento original y rectificado.
    With wsRecti
        If .AutoFilterMode Then .AutoFilterMode = False
        .Columns("A:B").Insert Shift:=xlToRight
        .Range("A1").Value = "Original"
        .Range("B1").Value = "Recti"

        ultimaRecti = UltimaFila(wsRecti, "C")

        If ultimaRecti < 2 Then Err.Raise vbObjectError + 104, , _
            "Recti no contiene registros."

        RellenarYFijar .Range("B2:B" & ultimaRecti), _
            "=RC[3]&""-""&RC[1]&""-""&RC[2]"

        RellenarYFijar .Range("A2:A" & ultimaRecti), _
            "=RC[9]&""-""&RC[8]&""-""&RC[7]"

        .Columns("A:B").AutoFit
    End With

    '2. PARTIDAS: llave común para sumar los valores por pedimento.
    With wsPartidas
        If .AutoFilterMode Then .AutoFilterMode = False
        .Columns("A:A").Insert Shift:=xlToRight
        .Range("A1").Value = "Llave"

        ultimaPartidas = UltimaFila(wsPartidas, "B")

        If ultimaPartidas < 2 Then Err.Raise vbObjectError + 105, , _
            "Partidas no contiene registros."

        RellenarYFijar .Range("A2:A" & ultimaPartidas), _
            "=RC[3]&""-""&RC[1]&""-""&RC[2]"

        .Columns("A:A").AutoFit
    End With

    '3. CASOS AF: copiar únicamente los registros con caso AF.
    With wsCasos
        If .AutoFilterMode Then .AutoFilterMode = False

        .Range("A1:H" & Application.Max(1, UltimaFila(wsCasos, "A"))) _
            .AutoFilter Field:=4, Criteria1:="AF"

        Set wsAF = wb.Worksheets.Add(After:=wsCasos)
        wsAF.Name = "Casos AF"

        .Range("A1:H" & Application.Max(1, UltimaFila(wsCasos, "A"))) _
            .SpecialCells(xlCellTypeVisible).Copy _
            Destination:=wsAF.Range("A1")

        .AutoFilterMode = False
    End With

    With wsAF
        .Columns("A:A").Insert Shift:=xlToRight
        .Range("A1").Value = "llave"
        ultimaAF = UltimaFila(wsAF, "B")

        If ultimaAF >= 2 Then
            RellenarYFijar .Range("A2:A" & ultimaAF), _
                "=RC[3]&""-""&RC[1]&""-""&RC[2]"
        End If

        .Columns("A:A").AutoFit
    End With

    '4. ENC: llave, limpieza de columnas y rectificaciones R1 a R5.
    With wsEnc
        If .AutoFilterMode Then .AutoFilterMode = False

        .Columns("A:A").Insert Shift:=xlToRight
        .Range("A1").Value = "llave"
        ultimaEnc = UltimaFila(wsEnc, "B")

        If ultimaEnc < 2 Then Err.Raise vbObjectError + 106, , _
            "Enc no contiene registros."

        RellenarYFijar .Range("A2:A" & ultimaEnc), _
            "=RC[3]&""-""&RC[1]&""-""&RC[2]"

        'Misma eliminación de columnas que en las macros originales.
        .Columns("G:AC").Delete Shift:=xlToLeft
        .Columns("H:H").Delete Shift:=xlToLeft

        For i = 1 To 5
            .Cells(1, 8 + i).Value = "R" & i
        Next i

        'R1 busca la llave original; R2:R5 buscan la rectificación anterior.
        RellenarYFijar .Range("I2:I" & ultimaEnc), _
            "=VLOOKUP(RC[-8],Recti!R2C1:R" & ultimaRecti & "C2,2,FALSE)"

        For i = 10 To 13
            RellenarYFijar .Range(.Cells(2, i), .Cells(ultimaEnc, i)), _
                "=VLOOKUP(RC[-1],Recti!R2C1:R" & ultimaRecti & "C2,2,FALSE)"
        Next i

        .Range("N1").Value = "Pedimento Correcto"
        RellenarYFijar .Range("N2:N" & ultimaEnc), _
            "=IFNA(IFNA(IFNA(IFNA(IFNA(RC[-1],RC[-2]),RC[-3]),RC[-4]),RC[-5]),RC[-13])"

        .Range("O1").Value = "AF"
        RellenarYFijar .Range("O2:O" & ultimaEnc), _
            "=VLOOKUP(RC[-1],'Casos AF'!R1C1:R" & _
            Application.Max(1, ultimaAF) & "C5,5,FALSE)"

        'P y Q se calculan por separado en cada hoja de salida.
        .Range("P1").Value = "Valor"
        .Range("Q1").Value = "Valor Agregado"
    End With

    '5. Crear importaciones y exportaciones desde la misma Enc preparada.
    CrearSalidaRetorno wb, wsEnc, wsPartidas, _
        ultimaEnc, ultimaPartidas, False

    CrearSalidaRetorno wb, wsEnc, wsPartidas, _
        ultimaEnc, ultimaPartidas, True

    wsEnc.AutoFilterMode = False
    Application.CutCopyMode = False

    MsgBox "Proceso terminado: se crearon Importaciones MP y Exportaciones MP.", _
           vbInformation

Salida:
    Application.Calculation = calculoAnterior
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ManejarError:
    MsgBox "No se completó el proceso: " & Err.Description & vbCrLf & _
           "Si ya se modificaron hojas, vuelve a empezar con una copia limpia.", _
           vbExclamation
    Resume Salida

End Sub

Private Sub CrearSalidaRetorno( _
    ByVal wb As Workbook, _
    ByVal wsEnc As Worksheet, _
    ByVal wsPartidas As Worksheet, _
    ByVal ultimaEnc As Long, _
    ByVal ultimaPartidas As Long, _
    ByVal esExportacion As Boolean)

    Dim wsSalida As Worksheet
    Dim rangoFiltro As Range
    Dim ultimaSalida As Long
    Dim columnaImporte As Long
    Dim tituloImporte As String
    Dim nombreSalida As String

    If esExportacion Then
        nombreSalida = "Exportaciones MP"
        tituloImporte = "Valor Comercial"
        columnaImporte = 11  'Partidas!K
    Else
        nombreSalida = "Importaciones MP"
        tituloImporte = "Valor Aduana"
        columnaImporte = 10  'Partidas!J
    End If

    'Calcular sobre Enc antes de filtrar y copiar.
    With wsEnc
        .Range("P1").Value = tituloImporte

        RellenarYFijar .Range("P2:P" & ultimaEnc), _
            "=SUMIF(Partidas!R2C1:R" & ultimaPartidas & "C1," & _
            "RC[-2],Partidas!R2C" & columnaImporte & _
            ":R" & ultimaPartidas & "C" & columnaImporte & ")"

        If esExportacion Then
            RellenarYFijar .Range("Q2:Q" & ultimaEnc), _
                "=SUMIF(Partidas!R2C1:R" & ultimaPartidas & "C1," & _
                "RC[-3],Partidas!R2C17:R" & ultimaPartidas & "C17)"
        Else
            .Range("Q2:Q" & ultimaEnc).ClearContents
        End If

        If .AutoFilterMode Then .AutoFilterMode = False
        Set rangoFiltro = .Range("A1:Q" & ultimaEnc)

        rangoFiltro.AutoFilter Field:=5, _
            Criteria1:=IIf(esExportacion, "2", "1")

        If esExportacion Then
            rangoFiltro.AutoFilter Field:=6, _
                Criteria1:=Array("V5", "RT", "V1"), _
                Operator:=xlFilterValues
        Else
            rangoFiltro.AutoFilter Field:=6, _
                Criteria1:="=IN", Operator:=xlOr, Criteria2:="=V1"
        End If

        rangoFiltro.AutoFilter Field:=7, Criteria1:="1"
        rangoFiltro.AutoFilter Field:=15, Criteria1:="#N/A"
    End With

    Set wsSalida = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsSalida.Name = nombreSalida

    'Se copia el encabezado y únicamente las filas visibles.
    wsEnc.Range("A1:Q" & ultimaEnc) _
        .SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsSalida.Range("A3")

    With wsSalida
        If .AutoFilterMode Then .AutoFilterMode = False

        'Quitar R1:R5. El importe de P pasa a K.
        .Columns("I:M").Delete Shift:=xlToLeft

        .Range("A1").Value = "Filtrar únicamente los últimos 12 meses"
        .Range("A1").Font.Color = vbRed

        ultimaSalida = UltimaFila(wsSalida, "A")
        .Range("K2").Formula = "=SUBTOTAL(9,K4:K" & _
            Application.Max(4, ultimaSalida) & ")"
        .Range("K2").NumberFormat = "#,##0.00"
        .Range("K2").Font.Bold = True
        .Range("K2").Font.Size = 12

        .Range("A3:L" & Application.Max(3, ultimaSalida)).AutoFilter
        .Columns("A:L").AutoFit
    End With

    wsEnc.AutoFilterMode = False
    Application.CutCopyMode = False

End Sub

Private Sub RellenarYFijar(ByVal destino As Range, ByVal formulaR1C1 As String)
    destino.FormulaR1C1 = formulaR1C1
    destino.Calculate
    destino.Value = destino.Value
End Sub

Private Function UltimaFila(ByVal ws As Worksheet, ByVal columna As String) As Long
    UltimaFila = ws.Cells(ws.Rows.Count, columna).End(xlUp).Row
End Function

Private Function ExisteHoja(ByVal wb As Workbook, ByVal nombre As String) As Boolean
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nombre, vbTextCompare) = 0 Then
            ExisteHoja = True
            Exit Function
        End If
    Next ws
End Function
