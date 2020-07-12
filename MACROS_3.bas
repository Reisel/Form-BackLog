Attribute VB_Name = "MACROS_3"
'VARIBLES SOLPED
Public AÑO, MES, CantActivos, CantInactivos, Total, GrpInactivo, AñoInactivo, MesInactivos, TipoSolped, _
    Texto, GrpArticulo, Solicitante As Long
Public UltimaFila, UltimaFila2, Vuelta As Long
Public RangoFechas, CeldaFechas As Range
Public i, A, B, C As Long
Public FechasMatriz() As Long
Public RangoFechas2, CeldaFechas2 As Range
'VARIABLES SOLPED INACTIVAS
'VARIABLES PROCESO ACTIVOS
Public ProcesosRango, ProcesosCelda As Range
Public UltimaFilaProcesos As Long
'VARIABLES RESUMEN PROCESOS
Public ReProcesosMatriz() As String
Public UltPDC As Long
'TOTAL SOLPED
Public TotalActivo, TotalInactivo, TotalSolped As Long

Sub Solped()
Application.Calculation = xlCalculationManual
Application.StatusBar = "Cargando Matriz de Solped Activa/Inactiva"

If ThisWorkbook.Worksheets("REPORTE").FilterMode = True Then
    Worksheets("REPORTE").ShowAllData
End If

Worksheets("MM-CO-PA-0002C").Activate
UltimaFila = Worksheets("MM-CO-PA-0002C").Range("I" & Rows.Count).End(xlUp).Row
Set RangoFechas = Worksheets("MM-CO-PA-0002C").Range(Cells(2, 9), Cells(UltimaFila, 9))
i = 0
A = 0
Vuelta = 0

For Each CeldaFechas In RangoFechas
    Vuelta = Vuelta + 1
    Application.StatusBar = "Cargando Matriz de Solped Activa/Inactiva " & Format((Vuelta / RangoFechas.Count) * 100, "0") & "%"
    
    MesFecha = Format(CeldaFechas, "MM")
    AñoFecha = Format(CeldaFechas, "YYYY")
    If i = 0 Then
        ReDim Preserve FechasMatriz(4, i)
        FechasMatriz(0, i) = AñoFecha
        FechasMatriz(1, i) = MesFecha
        If CeldaFechas.Offset(0, 19) = "Inactivos" Then
            FechasMatriz(2, i) = 0  'ACTIVOS
            FechasMatriz(3, i) = 1  'INACTIVOS
        Else
            FechasMatriz(2, i) = 1  'ACTIVOS
            FechasMatriz(3, i) = 0  'INACTIVOS
        End If
        GoTo Siguiente
    End If
    
        For B = 0 To A
        If FechasMatriz(0, B) = AñoFecha And FechasMatriz(1, B) = MesFecha Then
                If CeldaFechas.Offset(0, 19) = "Inactivos" Then
                    FechasMatriz(2, B) = FechasMatriz(2, B) + 0
                    FechasMatriz(3, B) = FechasMatriz(3, B) + 1
                Else
                    FechasMatriz(2, B) = FechasMatriz(2, B) + 1
                    FechasMatriz(3, B) = FechasMatriz(3, B) + 0
                End If
            GoTo YaExiste
        End If
        Next B
        
        ReDim Preserve FechasMatriz(4, i)
        FechasMatriz(0, i) = AñoFecha
        FechasMatriz(1, i) = MesFecha
            If CeldaFechas.Offset(0, 19) = "Inactivos" Then
                FechasMatriz(2, i) = 0  'ACTIVOS
                FechasMatriz(3, i) = 1  'INACTIVOS
            Else
                FechasMatriz(2, i) = 1  'ACTIVOS
                FechasMatriz(3, i) = 0  'INACTIVOS
            End If
Siguiente:
    A = i
    i = i + 1
YaExiste:

Next CeldaFechas

'*********************************************************************************
'PAGINA 2
Application.StatusBar = "Cargando Matriz de Solped Activa/Inactiva -Part 2-"

Worksheets("MM-CO-PA-0002C (2 PART)").Activate
UltimaFila2 = Worksheets("MM-CO-PA-0002C (2 PART)").Range("I" & Rows.Count).End(xlUp).Row
If UltimaFila2 < 2 Then
    GoTo NOHAYPAGINA2
End If

Set RangoFechas2 = Worksheets("MM-CO-PA-0002C (2 PART)").Range(Cells(2, 9), Cells(UltimaFila2, 9))
Vuelta = 0
For Each CeldaFechas2 In RangoFechas2
    Vuelta = Vuelta + 1
    Application.StatusBar = "Cargando Matriz de Solped Activa/Inactiva -Part 2-" & Format((Vuelta / RangoFechas2.Count) * 100, "0") & "%"
    
    MesFecha = Format(CeldaFechas2, "MM")
    AñoFecha = Format(CeldaFechas2, "YYYY")
    If i = 0 Then
        ReDim Preserve FechasMatriz(4, i)
        FechasMatriz(0, i) = AñoFecha
        FechasMatriz(1, i) = MesFecha
        If CeldaFechas2.Offset(0, 19) = "Inactivos" Then
            FechasMatriz(2, i) = 0  'ACTIVOS
            FechasMatriz(3, i) = 1  'INACTIVOS
        Else
            FechasMatriz(2, i) = 1  'ACTIVOS
            FechasMatriz(3, i) = 0  'INACTIVOS
        End If
        GoTo Siguiente22
    End If
    
        For B = 0 To A
        If FechasMatriz(0, B) = AñoFecha And FechasMatriz(1, B) = MesFecha Then
                If CeldaFechas2.Offset(0, 19) = "Inactivos" Then
                    FechasMatriz(2, B) = FechasMatriz(2, B) + 0
                    FechasMatriz(3, B) = FechasMatriz(3, B) + 1
                Else
                    FechasMatriz(2, B) = FechasMatriz(2, B) + 1
                    FechasMatriz(3, B) = FechasMatriz(3, B) + 0
                End If
            GoTo YaExiste22
        End If
        Next B
        
        ReDim Preserve FechasMatriz(4, i)
        FechasMatriz(0, i) = AñoFecha
        FechasMatriz(1, i) = MesFecha
            If CeldaFechas2.Offset(0, 19) = "Inactivos" Then
                FechasMatriz(2, i) = 0  'ACTIVOS
                FechasMatriz(3, i) = 1  'INACTIVOS
            Else
                FechasMatriz(2, i) = 1  'ACTIVOS
                FechasMatriz(3, i) = 0  'INACTIVOS
            End If
Siguiente22:
    A = i
    i = i + 1
YaExiste22:

Next CeldaFechas2
NOHAYPAGINA2:
'FIN PAGINA 2
'*********************************************************************************
Application.StatusBar = "Descargando Matriz de Solped Activa/Inactiva"
'CALCULO TOTAL
Worksheets("REPORTE").Activate
Worksheets("REPORTE").Cells(3, 20).Select
TotalActivo = 0
TotalInactivo = 0
Total = 0

For B = 0 To A
    Application.StatusBar = "Descargando Matriz de Solped Activa/Inactiva " & Format((B / A) * 100, "0") & "%"

    FechasMatriz(4, B) = FechasMatriz(2, B) + FechasMatriz(3, B)
    'DESCARGA DE MATRIZ
    Worksheets("REPORTE").Cells(3 + B, 20) = FechasMatriz(0, B)
    Worksheets("REPORTE").Cells(3 + B, 21) = FechasMatriz(1, B)
    Worksheets("REPORTE").Cells(3 + B, 22) = FechasMatriz(2, B)
    Worksheets("REPORTE").Cells(3 + B, 22).NumberFormat = "#,##0"
    Worksheets("REPORTE").Cells(3 + B, 23) = FechasMatriz(3, B)
    Worksheets("REPORTE").Cells(3 + B, 23).NumberFormat = "#,###"
    Worksheets("REPORTE").Cells(3 + B, 24) = FechasMatriz(4, B)
    Worksheets("REPORTE").Cells(3 + B, 24).NumberFormat = "#,###"
    TotalActivo = TotalActivo + FechasMatriz(2, B)
    TotalInactivo = TotalInactivo + FechasMatriz(3, B)
    TotalSolped = TotalActivo + TotalInactivo
Next B

'FORMATO DEL CUADRO
Worksheets("REPORTE").Range(Cells(3, 20), Cells(A + 3, 24)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
Worksheets("REPORTE").Cells(A + 4, 22) = TotalActivo
Worksheets("REPORTE").Cells(A + 4, 22).NumberFormat = "#,##0"
Worksheets("REPORTE").Cells(A + 4, 23) = TotalInactivo
Worksheets("REPORTE").Cells(A + 4, 23).NumberFormat = "#,##0"
Worksheets("REPORTE").Cells(A + 4, 24) = TotalSolped
Worksheets("REPORTE").Cells(A + 4, 24).NumberFormat = "#,##0"

Worksheets("REPORTE").Cells(A + 4, 21).FormulaLocal = "Total"
Worksheets("REPORTE").Cells(A + 4, 21).Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
    End With

'ORDENAR CUADRO
Application.StatusBar = "Ordentando Solped Activa/Inactiva"
Worksheets("REPORTE").Range(Cells(2, 20), Cells(3 + A, 24)).Select
    Selection.Sort Key1:=Cells(3, 20), Order1:=xlDescending, Key2:=Cells(3, 21) _
        , Order2:=xlDescending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal

Application.StatusBar = False

Call SolpedInactivas

End Sub

Sub SolpedInactivas()
Application.Calculation = xlCalculationManual
Application.StatusBar = "Cargando Cuadro de Solped´s de Grp Com. Inactivas"

A = 0
Worksheets("REPORTE").Cells(3, 26).Select
Vuelta = 0
For Each CeldaFechas In RangoFechas
Vuelta = Vuelta + 1
Application.StatusBar = "Cargando Cuadro de Solped´s Inactivas " & Format((Vuelta / RangoFechas.Count) * 100, "0") & "%"

    If CeldaFechas.Offset(0, 19) = "Inactivos" Then
        Worksheets("REPORTE").Cells(3 + A, 26) = CeldaFechas.Offset(0, 3) 'GRUPO COMPRA
        Worksheets("REPORTE").Cells(3 + A, 27) = Format(CeldaFechas, "YYYY")    'AÑO
        Worksheets("REPORTE").Cells(3 + A, 28) = Format(CeldaFechas, "MM")  'MES
        Worksheets("REPORTE").Cells(3 + A, 29) = CeldaFechas.Offset(0, 7)   'Cdoc
        Worksheets("REPORTE").Cells(3 + A, 30) = CeldaFechas.Offset(0, 2)   'I
        Worksheets("REPORTE").Cells(3 + A, 31) = CeldaFechas.Offset(0, -6)   'Solped
        Worksheets("REPORTE").Cells(3 + A, 32) = CeldaFechas.Offset(0, -5) * 1 'Pos
        Worksheets("REPORTE").Cells(3 + A, 33) = CeldaFechas.Offset(0, -4)   'Codigo
        Worksheets("REPORTE").Cells(3 + A, 34) = CeldaFechas.Offset(0, -3)  'Texto
        Worksheets("REPORTE").Cells(3 + A, 35) = CeldaFechas.Offset(0, 10)  'Grp Art
        Worksheets("REPORTE").Cells(3 + A, 36) = CeldaFechas.Offset(0, 13)  'Creado Por
        A = A + 1
    End If
Next CeldaFechas

'*****************************************
'INICIO PART 2
Application.StatusBar = "Cargando Cuadro de Solped´s de Grp Com. Inactivas -Part 2-"

If UltimaFila2 < 2 Then
    GoTo NOHAYPAGINA3
End If

Vuelta = 0
For Each CeldaFechas2 In RangoFechas2
Vuelta = Vuelta + 1
Application.StatusBar = "Cargando Cuadro de Solped´s de Grp Com. Inactivas -Part 2-" & Format((Vuelta / RangoFechas.Count) * 100, "0") & "%"

    If CeldaFechas2.Offset(0, 19) = "Inactivos" Then
        Worksheets("REPORTE").Cells(3 + A, 26) = CeldaFechas2.Offset(0, 3) 'GRUPO COMPRA
        Worksheets("REPORTE").Cells(3 + A, 27) = Format(CeldaFechas2, "YYYY")    'AÑO
        Worksheets("REPORTE").Cells(3 + A, 28) = Format(CeldaFechas2, "MM")  'MES
        Worksheets("REPORTE").Cells(3 + A, 29) = CeldaFechas2.Offset(0, 7)   'Cdoc
        Worksheets("REPORTE").Cells(3 + A, 30) = CeldaFechas2.Offset(0, 2)   'I
        Worksheets("REPORTE").Cells(3 + A, 31) = CeldaFechas2.Offset(0, -6)   'Solped
        Worksheets("REPORTE").Cells(3 + A, 32) = CeldaFechas2.Offset(0, -5) * 1 'Pos
        Worksheets("REPORTE").Cells(3 + A, 33) = CeldaFechas2.Offset(0, -4)   'Codigo
        Worksheets("REPORTE").Cells(3 + A, 34) = CeldaFechas2.Offset(0, -3)  'Texto
        Worksheets("REPORTE").Cells(3 + A, 35) = CeldaFechas2.Offset(0, 10)  'Grp Art
        Worksheets("REPORTE").Cells(3 + A, 36) = CeldaFechas2.Offset(0, 13)  'Creado Por
        A = A + 1
    End If
Next CeldaFechas2

NOHAYPAGINA3:
'FIN PART 2
'****************************************

'ORDENAR CUADRO
Application.StatusBar = "Ordenando Solped´s de Grp Com. Inactivas"
Worksheets("REPORTE").Range(Cells(2, 26), Cells(3 + A, 36)).Select
    Selection.Sort Key1:=Cells(3, 27), Order1:=xlDescending, Key2:=Cells(3, 28) _
        , Order2:=xlDescending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal


'IDENTIFICANDO AREAS
Application.Calculation = xlCalculationManual
Application.StatusBar = "Ordenando Solped´s Inactivas"
UltimaFilaSolped = Worksheets("REPORTE").Range("Z" & Rows.Count).End(xlUp).Row
'Departamento
Worksheets("REPORTE").Cells(3, 37).Select
ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(FC(-1);Usuarios!C(-36):C(-35);2;FALSO));"""";BUSCARV(FC(-1);Usuarios!C(-36):C(-35);2;FALSO))"
Selection.Copy
Range(Cells(3, 37), Cells(UltimaFilaSolped, 37)).Select
ActiveSheet.Paste
Application.CutCopyMode = False
'Area
Worksheets("REPORTE").Cells(3, 38).Select
ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(FC(-1);Usuarios!C(-36):C(-35);2;FALSO));"""";BUSCARV(FC(-1);Usuarios!C(-36):C(-35);2;FALSO))"
Selection.Copy
Range(Cells(3, 38), Cells(UltimaFilaSolped, 38)).Select
ActiveSheet.Paste
Application.CutCopyMode = False
'CALCULAR
Worksheets("MM-CO-PA-0002C").Calculate
Application.Calculation = xlCalculationAutomatic
'PEGAR VALORES
Range(Cells(3, 37), Cells(UltimaFilaSolped, 38)).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
Application.CutCopyMode = False

Worksheets("REPORTE").Cells(3, 26).Select

Application.StatusBar = False

Call Procesos

Worksheets("REPORTE").Cells(3, 26).Select

End Sub

Sub Procesos()
Application.Calculation = xlCalculationManual
Application.StatusBar = "Cargando Cuadro de Procesos Activos"

Worksheets("PET (MM-CO-PA-0004)").Activate
UltimaFilaProcesos = Worksheets("PET (MM-CO-PA-0004)").Range("S" & Rows.Count).End(xlUp).Row
Set ProcesosRango = Worksheets("PET (MM-CO-PA-0004)").Range(Cells(2, 19), Cells(UltimaFilaProcesos, 19))

A = 0
Worksheets("REPORTE").Activate
Worksheets("REPORTE").Cells(3, 40).Select
Vuelta = 0
For Each ProcesosCelda In ProcesosRango
Vuelta = Vuelta + 1
Application.StatusBar = "Cargando Cuadro de Procesos Activos " & Format((Vuelta / ProcesosRango.Count) * 100, "0") & "%"

    If ProcesosCelda = "" And ProcesosCelda.Offset(0, -3) <> "B" And ProcesosCelda.Offset(0, 1) = "" Then
        Worksheets("REPORTE").Cells(3 + A, 40) = ProcesosCelda.Offset(0, -10) 'GRUPO COMPRA
        Worksheets("REPORTE").Cells(3 + A, 41) = ProcesosCelda.Offset(0, -9) 'Nombre
        Worksheets("REPORTE").Cells(3 + A, 42) = Format(ProcesosCelda.Offset(0, -12), "YYYY")    'AÑO
        Worksheets("REPORTE").Cells(3 + A, 43) = Format(ProcesosCelda.Offset(0, -12), "MM")  'MES
        Worksheets("REPORTE").Cells(3 + A, 44) = ProcesosCelda.Offset(0, -15)   'Proceso
        Worksheets("REPORTE").Cells(3 + A, 45) = ProcesosCelda.Offset(0, -11)   'Tipo Doc
        Worksheets("REPORTE").Cells(3 + A, 46) = ProcesosCelda.Offset(0, -17)   'Solped
        Worksheets("REPORTE").Cells(3 + A, 47) = ProcesosCelda.Offset(0, -16) * 1 'Pos
        Worksheets("REPORTE").Cells(3 + A, 48) = ProcesosCelda.Offset(0, -7)   'Codigo
        Worksheets("REPORTE").Cells(3 + A, 49) = ProcesosCelda.Offset(0, -5)  'Texto
        A = A + 1
    End If
Next ProcesosCelda

'ORDENAR CUADRO
Application.StatusBar = "Ordenando Procesos Activos"
Worksheets("REPORTE").Range(Cells(2, 40), Cells(3 + A, 49)).Select
    Selection.Sort Key1:=Cells(3, 42), Order1:=xlAscending, Key2:=Cells(3, 43) _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal

Application.StatusBar = False

Call ResumenProcesos

End Sub

Sub ResumenProcesos()


Application.Calculation = xlCalculationManual
Application.StatusBar = "Cargando Matriz Resumen Procesos"

i = 0
A = 0
Vuelta = 0
Worksheets("REPORTE").Activate

For Each ProcesosCelda In ProcesosRango
    Vuelta = Vuelta + 1
    Application.StatusBar = "Cargando Matriz Resumen Procesos " & Format((Vuelta / RangoFechas.Count) * 100, "0") & "%"
    
    If ProcesosCelda = "" And ProcesosCelda.Offset(0, -3) <> "B" And ProcesosCelda.Offset(0, 1) = "" Then
    
        'PRIMERA VUELTA
        If i = 0 Then
            ReDim Preserve ReProcesosMatriz(5, i)
            ReProcesosMatriz(0, i) = ProcesosCelda.Offset(0, -10) 'Grupo de Compra
            ReProcesosMatriz(1, i) = ProcesosCelda.Offset(0, -9)   'Nombre
            ReProcesosMatriz(2, i) = Format(ProcesosCelda.Offset(0, -12), "YYYY")   'Año
            ReProcesosMatriz(3, i) = Format(ProcesosCelda.Offset(0, -12), "MM")   'Mes
            ReProcesosMatriz(4, i) = ProcesosCelda.Offset(0, -15)   'Proceso
            ReProcesosMatriz(5, i) = ProcesosCelda.Offset(0, -11)   'Tipo
            GoTo Siguiente
        End If
        'VERIFICAR SI YA EXISTE
        For B = 0 To A
            If ReProcesosMatriz(4, B) = ProcesosCelda.Offset(0, -15) Then
                GoTo YaExiste
            End If
        Next B
        'NUEVOS NUMEROS
        ReDim Preserve ReProcesosMatriz(5, i)
        ReProcesosMatriz(0, i) = ProcesosCelda.Offset(0, -10) 'Grupo de Compra
        ReProcesosMatriz(1, i) = ProcesosCelda.Offset(0, -9)   'Nombre
        ReProcesosMatriz(2, i) = Format(ProcesosCelda.Offset(0, -12), "YYYY")   'Año
        ReProcesosMatriz(3, i) = Format(ProcesosCelda.Offset(0, -12), "MM")   'Mes
        ReProcesosMatriz(4, i) = ProcesosCelda.Offset(0, -15)   'Proceso
        ReProcesosMatriz(5, i) = ProcesosCelda.Offset(0, -11)   'Tipo
        
Siguiente:
    A = i
    i = i + 1
YaExiste:
    End If

Next ProcesosCelda

Worksheets("REPORTE").Activate
Worksheets("REPORTE").Cells(3, 51).Select
For B = 0 To A
    Application.StatusBar = "Descargando Matriz Resumen Procesos " & Format((B / A) * 100, "0") & "%"
    'DESCARGA DE MATRIZ
    Worksheets("REPORTE").Cells(3 + B, 51) = ReProcesosMatriz(0, B)
    Worksheets("REPORTE").Cells(3 + B, 52) = ReProcesosMatriz(1, B)
    Worksheets("REPORTE").Cells(3 + B, 53) = ReProcesosMatriz(2, B)
    Worksheets("REPORTE").Cells(3 + B, 54) = ReProcesosMatriz(3, B)
    Worksheets("REPORTE").Cells(3 + B, 55) = ReProcesosMatriz(4, B)
    Worksheets("REPORTE").Cells(3 + B, 56) = ReProcesosMatriz(5, B)
Next B



'ORDENAR HOJA PDC
Worksheets("PDC").Activate
UltPDC = Worksheets("PDC").Range("A" & Rows.Count).End(xlUp).Row
Application.StatusBar = "Ordenando PDC"
Worksheets("PDC").Range(Cells(1, 1), Cells(UltPDC, 30)).Select
    Selection.Sort Key1:=Cells(2, 16), Order1:=xlDescending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

Worksheets("REPORTE").Activate

'Nombre
Worksheets("REPORTE").Cells(3, 57).Select
ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(BC3;PDC!D:G;4;FALSO));""NO EN PUNTO DE CUENTA"";EXTRAE(BUSCARV(BC3;PDC!D:G;4;FALSO);HALLAR(""ADQUISICI"";BUSCARV(BC3;PDC!D:G;4;FALSO))+15;99))"
Selection.Copy
Range(Cells(3, 57), Cells(3 + A, 57)).Select
ActiveSheet.Paste
Application.CutCopyMode = False

'Observación
Worksheets("REPORTE").Cells(3, 58).Select
ActiveCell.FormulaLocal = "=SI(BE3=""NO EN PUNTO DE CUENTA"";"""";BUSCARV(BC3;PDC!D:O;12;FALSO))"
Selection.Copy
Range(Cells(3, 58), Cells(3 + A, 58)).Select
ActiveSheet.Paste
Application.CutCopyMode = False

'Status Final
Worksheets("REPORTE").Cells(3, 59).Select
ActiveCell.FormulaLocal = "=SI(BE3=""NO EN PUNTO DE CUENTA"";"""";BUSCARV(BC3;PDC!D:R;15;FALSO))"
Selection.Copy
Range(Cells(3, 59), Cells(3 + A, 59)).Select
ActiveSheet.Paste
Application.CutCopyMode = False

'Fecha
Worksheets("REPORTE").Cells(3, 60).Select
ActiveCell.FormulaLocal = "=SI(BE3=""NO EN PUNTO DE CUENTA"";"""";BUSCARV(BC3;PDC!D:P;13;FALSO))"
ActiveCell.NumberFormat = "DD/MM/YYYY"
Selection.Copy
Range(Cells(3, 60), Cells(3 + A, 60)).Select
ActiveSheet.Paste
Application.CutCopyMode = False


'CALCULAR
Worksheets("REPORTE").Calculate
Application.Calculation = xlCalculationAutomatic
'PEGAR VALORES
Worksheets("REPORTE").Range(Cells(3, 57), Cells(3 + A, 60)).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
Application.CutCopyMode = False

'ORDENAR CUADRO
Application.StatusBar = "Ordenando Resumen Procesos"
Worksheets("REPORTE").Range(Cells(2, 51), Cells(3 + A, 60)).Select
    Selection.Sort Key1:=Cells(3, 53), Order1:=xlAscending, Key2:=Cells(3, 54) _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal

Application.StatusBar = False

End Sub
