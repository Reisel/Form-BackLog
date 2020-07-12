Attribute VB_Name = "MACROS_2"
Option Explicit

Sub Buscar()
Attribute Buscar.VB_Description = "Macro grabada el 31/08/2018 por Administrador"
Attribute Buscar.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Rango As Range
Dim Celda, CeldaProceso As Range
Dim CeldaT As Long
Dim GrupoCompra As String
Dim SolpedBandeja As Long
Dim PosicionBandeja As Long
Dim RangoBandeja As Range
Dim BandejaGrupoCompra As Range
Dim RangoBandeja2 As Range
Dim BandejaGrupoCompra2 As Range
Dim FinalBandeja As Long
Dim SolpedProceso As Long
Dim RangoProceso As Range
Dim SolpedProcesoCelda As Range
Dim FinalProceso As Long
Dim MensajeTexto, MensajeTextoBandeja, MensajeTextoProceso, GrupoCompraSuplente As String
Dim SolpedAnterior As Long
Dim ProcesoGrupoCompra As Range
Dim SolpedBandejaCelda As Range
Dim PosicionProceso As Long
Dim FilaGrupo, FilaGrupoProceso, A, B, C, D, T, Z, X As Long
Dim NombreComprador, NombreSuplente As String
Dim RefYaTieneNombre As String
Dim RefConteo2 As Range
Dim Contar, i, e As Long
Dim Repetido(1000) As String
Dim UltimaFilaMonitorio, Vuelta As Long

Application.Calculation = xlCalculationManual

Worksheets("REPORTE").Activate

If ThisWorkbook.Worksheets("MM-CO-PA-0002C").FilterMode = True Then
    Worksheets("MM-CO-PA-0002C").ShowAllData
End If
If ThisWorkbook.Worksheets("PET (MM-CO-PA-0004)").FilterMode = True Then
    Worksheets("PET (MM-CO-PA-0004)").ShowAllData
End If
If ThisWorkbook.Worksheets("REPORTE").FilterMode = True Then
        Worksheets("REPORTE").ShowAllData
End If

A = Worksheets("REPORTE").Range("A" & Rows.Count).End(xlUp).Row
B = Worksheets("REPORTE").Range("H" & Rows.Count).End(xlUp).Row
D = Worksheets("REPORTE").Range("O" & Rows.Count).End(xlUp).Row
T = Worksheets("REPORTE").Range("T" & Rows.Count).End(xlUp).Row
Z = Worksheets("REPORTE").Range("Z" & Rows.Count).End(xlUp).Row
X = Worksheets("REPORTE").Range("AN" & Rows.Count).End(xlUp).Row
C = Application.WorksheetFunction.Max(A, B, D, T, Z, X)

If C < 3 Then
    C = 3
End If

Worksheets("REPORTE").Range("A3:BH" & C).Clear

FilaGrupo = 2
UltimaFilaMonitorio = Worksheets("Monitoreo").Range("Z" & Rows.Count).End(xlUp).Row
Set Rango = Worksheets("Monitoreo").Range("Z7:Z" & UltimaFilaMonitorio)

Worksheets("REPORTE").Cells(3, 1).Select
Vuelta = 0
For Each Celda In Rango
Vuelta = Vuelta + 1

Application.StatusBar = "Revisión de Tabla Monitoreo " & Format((Vuelta / Rango.Count) * 100, "0") & "%"

If IsNumeric(Celda.Offset(0, -21)) Then
    CeldaT = Celda.Offset(0, -21)
Else
    CeldaT = 0
End If

If CeldaT <> Celda.Value Then  'Tienen diferencia en cantidad Tratada  And (Celda.Offset(0, -22) = "QUIMICOS" Or Celda.Offset(0, -22) = "PARADAS" Or Celda.Offset(0, -22) = "RUTINA" Or Celda.Offset(0, -22) = "CORPORATIVO")
    '***PARA GRUPO DE INACTIVOS***
    If Celda.Offset(0, -24) = "Analista Exterior" Or Celda.Offset(0, -24) = "Inactivos" Then
    Dim UltimaFilaRef, FinalBandeja2 As Long
    Dim InactivosRang, InactivoCelda As Range
    Application.ScreenUpdating = False

    UltimaFilaRef = Worksheets("Ref").Range("B" & Rows.Count).End(xlUp).Row
    Set InactivosRang = Worksheets("Ref").Range("B3:B" & UltimaFilaRef)
    For Each InactivoCelda In InactivosRang
    
        If InactivoCelda = Celda.Offset(0, -24) Then
            GrupoCompra = InactivoCelda.Offset(0, -1)
            NombreComprador = InactivoCelda
            FinalBandeja = Worksheets("MM-CO-PA-0002C").Range("C" & Rows.Count).End(xlUp).Row
    
            Set RangoBandeja = Worksheets("MM-CO-PA-0002C").Range("L2:L" & FinalBandeja)
            For Each BandejaGrupoCompra In RangoBandeja 'Grupo de Compra
            
            If BandejaGrupoCompra.Value = GrupoCompra And BandejaGrupoCompra.Offset(0, 2) = "A" Then 'Si el grupo de compra es diferente y estatus es A
                SolpedBandeja = BandejaGrupoCompra.Offset(0, -9).Value 'Solped
                PosicionBandeja = BandejaGrupoCompra.Offset(0, -8) 'Posicion
                FinalProceso = Worksheets("PET (MM-CO-PA-0004)").Range("I" & Rows.Count).End(xlUp).Row
                
                Set RangoProceso = Worksheets("PET (MM-CO-PA-0004)").Range("I2:I" & FinalProceso)
                For Each ProcesoGrupoCompra In RangoProceso
                    PosicionProceso = ProcesoGrupoCompra.Offset(0, -6)
                    SolpedProceso = ProcesoGrupoCompra.Offset(0, -7)
                    
                    If ProcesoGrupoCompra.Value = GrupoCompra And ProcesoGrupoCompra.Offset(0, 7) = "A" And SolpedBandeja = SolpedProceso And PosicionBandeja = PosicionProceso And ProcesoGrupoCompra.Offset(0, 10) = "" Then 'Si el grupo de compra es diferente y estatus es A
                        GoTo SeEncontroSolped22
                    End If
                Next ProcesoGrupoCompra
                'LO QUE NO SE ENCONTRO
                RefYaTieneNombre = 0 'Para evitar repetición
                
                For Each ProcesoGrupoCompra In RangoProceso
                    
                    SolpedProceso = ProcesoGrupoCompra.Offset(0, -7)
                    PosicionProceso = ProcesoGrupoCompra.Offset(0, -6)
                    
                    If ProcesoGrupoCompra.Offset(0, 7) = "A" And SolpedBandeja = SolpedProceso And PosicionBandeja = PosicionProceso And ProcesoGrupoCompra.Offset(0, 10) = "" Then 'Si el grupo de compra es diferente y estatus es A
                        GrupoCompraSuplente = ProcesoGrupoCompra.Value
                        NombreSuplente = ProcesoGrupoCompra.Offset(0, 1)
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 1) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 2) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 3) = SolpedBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 4) = PosicionBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 5) = GrupoCompraSuplente
                        Worksheets("REPORTE").Cells(FilaGrupo, 6) = NombreSuplente
                        RefYaTieneNombre = 1
                    End If

                Next ProcesoGrupoCompra
                    
                    If RefYaTieneNombre = 0 Then
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 1) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 2) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 3) = SolpedBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 4) = PosicionBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 5) = "--"
                        Worksheets("REPORTE").Cells(FilaGrupo, 6) = "PETICIÓN ACTIVA" ' La solped esta en estatus A pero no se encuentra en la hoja MM-CO-PA-0004
                    End If
                MensajeTextoBandeja = MensajeTextoBandeja & vbNewLine & GrupoCompra & " " & SolpedBandeja & " " & PosicionBandeja
            End If
        
SeEncontroSolped22:
     
            Next BandejaGrupoCompra

            'GRUPO INACTIVO***INICO MODIFICACION PARA MM-CO-PA-0002C (2 PART)****************
            FinalBandeja = Worksheets("MM-CO-PA-0002C (2 PART)").Range("C" & Rows.Count).End(xlUp).Row
            If FinalBandeja < 2 Then GoTo CANCELAR2
    
            Set RangoBandeja = Worksheets("MM-CO-PA-0002C (2 PART)").Range("L2:L" & FinalBandeja)
            For Each BandejaGrupoCompra In RangoBandeja 'Grupo de Compra
                If BandejaGrupoCompra.Value = GrupoCompra And BandejaGrupoCompra.Offset(0, 2) = "A" Then 'Si el grupo de compra es diferente y estatus es A
                SolpedBandeja = BandejaGrupoCompra.Offset(0, -9).Value 'Solped
                PosicionBandeja = BandejaGrupoCompra.Offset(0, -8) 'Posicion
                FinalProceso = Worksheets("PET (MM-CO-PA-0004)").Range("I" & Rows.Count).End(xlUp).Row
                
                Set RangoProceso = Worksheets("PET (MM-CO-PA-0004)").Range("I2:I" & FinalProceso)
                    For Each ProcesoGrupoCompra In RangoProceso
                    PosicionProceso = ProcesoGrupoCompra.Offset(0, -6)
                    SolpedProceso = ProcesoGrupoCompra.Offset(0, -7)
                    If ProcesoGrupoCompra.Value = GrupoCompra And ProcesoGrupoCompra.Offset(0, 7) = "A" And SolpedBandeja = SolpedProceso And PosicionBandeja = PosicionProceso And ProcesoGrupoCompra.Offset(0, 10) = "" Then 'Si el grupo de compra es diferente y estatus es A
                        GoTo SeEncontroSolped23
                    End If
                Next ProcesoGrupoCompra
                'LO QUE NO SE ENCONTRO
                RefYaTieneNombre = 0 'Para evitar repetición
                
                For Each ProcesoGrupoCompra In RangoProceso
                    SolpedProceso = ProcesoGrupoCompra.Offset(0, -7)
                    PosicionProceso = ProcesoGrupoCompra.Offset(0, -6)
                    
                    If ProcesoGrupoCompra.Offset(0, 7) = "A" And SolpedBandeja = SolpedProceso And PosicionBandeja = PosicionProceso And ProcesoGrupoCompra.Offset(0, 10) = "" Then 'Si el grupo de compra es diferente y estatus es A
                        GrupoCompraSuplente = ProcesoGrupoCompra.Value
                        NombreSuplente = ProcesoGrupoCompra.Offset(0, 1)
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 1) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 2) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 3) = SolpedBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 4) = PosicionBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 5) = GrupoCompraSuplente
                        Worksheets("REPORTE").Cells(FilaGrupo, 6) = NombreSuplente
                        RefYaTieneNombre = 1
                    End If

                Next ProcesoGrupoCompra
                    If RefYaTieneNombre = 0 Then
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 1) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 2) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 3) = SolpedBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 4) = PosicionBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 5) = "--"
                        Worksheets("REPORTE").Cells(FilaGrupo, 6) = "PETICIÓN ACTIVA" ' La solped esta en estatus A pero no se encuentra en la hoja MM-CO-PA-0004
                    End If
            MensajeTextoBandeja = MensajeTextoBandeja & vbNewLine & GrupoCompra & " " & SolpedBandeja & " " & PosicionBandeja
            End If
        
SeEncontroSolped23:
     
            Next BandejaGrupoCompra
CANCELAR2:
'***PARA GRUPO DE INACTIVOS FIN*****FIN MODIFICACION PARA MM-CO-PA-0002C (2 PART)****************
    End If
    
    Next InactivoCelda
'***PARA GRUPO DE INACTIVOS FIN***

    End If
 
    
    GrupoCompra = Celda.Offset(0, -25)
    NombreComprador = Celda.Offset(0, -24)
    FinalBandeja = Worksheets("MM-CO-PA-0002C").Range("C" & Rows.Count).End(xlUp).Row
    
    Set RangoBandeja = Worksheets("MM-CO-PA-0002C").Range("L2:L" & FinalBandeja)
    For Each BandejaGrupoCompra In RangoBandeja 'Grupo de Compra
          If BandejaGrupoCompra.Value = GrupoCompra And BandejaGrupoCompra.Offset(0, 2) = "A" Then 'Si el grupo de compra es diferente y estatus es A
            SolpedBandeja = BandejaGrupoCompra.Offset(0, -9).Value 'Solped
            PosicionBandeja = BandejaGrupoCompra.Offset(0, -8) 'Posicion
            FinalProceso = Worksheets("PET (MM-CO-PA-0004)").Range("I" & Rows.Count).End(xlUp).Row
            Set RangoProceso = Worksheets("PET (MM-CO-PA-0004)").Range("I2:I" & FinalProceso)
            For Each ProcesoGrupoCompra In RangoProceso
                PosicionProceso = ProcesoGrupoCompra.Offset(0, -6)
                SolpedProceso = ProcesoGrupoCompra.Offset(0, -7)
                If ProcesoGrupoCompra.Value = GrupoCompra And ProcesoGrupoCompra.Offset(0, 7) = "A" And SolpedBandeja = SolpedProceso And PosicionBandeja = PosicionProceso And ProcesoGrupoCompra.Offset(0, 10) = "" Then 'Si el grupo de compra es diferente y estatus es A
                    GoTo SeEncontroSolped
                End If
            Next ProcesoGrupoCompra
            'LO QUE NO SE ENCONTRO
            RefYaTieneNombre = 0 'Para evitar repetición
                For Each ProcesoGrupoCompra In RangoProceso
                    SolpedProceso = ProcesoGrupoCompra.Offset(0, -7)
                    PosicionProceso = ProcesoGrupoCompra.Offset(0, -6)
                    If ProcesoGrupoCompra.Offset(0, 7) = "A" And SolpedBandeja = SolpedProceso And PosicionBandeja = PosicionProceso And ProcesoGrupoCompra.Offset(0, 10) = "" Then 'Si el grupo de compra es diferente y estatus es A
                        GrupoCompraSuplente = ProcesoGrupoCompra.Value
                        NombreSuplente = ProcesoGrupoCompra.Offset(0, 1)
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 1) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 2) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 3) = SolpedBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 4) = PosicionBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 5) = GrupoCompraSuplente
                        Worksheets("REPORTE").Cells(FilaGrupo, 6) = NombreSuplente
                        RefYaTieneNombre = 1
                    End If

                Next ProcesoGrupoCompra
                    If RefYaTieneNombre = 0 Then
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 1) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 2) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 3) = SolpedBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 4) = PosicionBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 5) = "--"
                        Worksheets("REPORTE").Cells(FilaGrupo, 6) = "PETICIÓN ACTIVA" ' La solped esta en estatus A pero no se encuentra en la hoja MM-CO-PA-0004
                    End If
            MensajeTextoBandeja = MensajeTextoBandeja & vbNewLine & GrupoCompra & " " & SolpedBandeja & " " & PosicionBandeja
        End If
        
SeEncontroSolped:
     
     Next BandejaGrupoCompra

'*************************************INICO MODIFICACION PARA MM-CO-PA-0002C (2 PART)****************
    FinalBandeja = Worksheets("MM-CO-PA-0002C (2 PART)").Range("C" & Rows.Count).End(xlUp).Row
    If FinalBandeja < 2 Then GoTo CANCELAR
    
    Set RangoBandeja = Worksheets("MM-CO-PA-0002C (2 PART)").Range("L2:L" & FinalBandeja)
    For Each BandejaGrupoCompra In RangoBandeja 'Grupo de Compra
        If BandejaGrupoCompra.Value = GrupoCompra And BandejaGrupoCompra.Offset(0, 2) = "A" Then 'Si el grupo de compra es diferente y estatus es A
            SolpedBandeja = BandejaGrupoCompra.Offset(0, -9).Value 'Solped
            PosicionBandeja = BandejaGrupoCompra.Offset(0, -8) 'Posicion
            FinalProceso = Worksheets("PET (MM-CO-PA-0004)").Range("I" & Rows.Count).End(xlUp).Row
            Set RangoProceso = Worksheets("PET (MM-CO-PA-0004)").Range("I2:I" & FinalProceso)
            For Each ProcesoGrupoCompra In RangoProceso
                PosicionProceso = ProcesoGrupoCompra.Offset(0, -6)
                SolpedProceso = ProcesoGrupoCompra.Offset(0, -7)
                If ProcesoGrupoCompra.Value = GrupoCompra And ProcesoGrupoCompra.Offset(0, 7) = "A" And SolpedBandeja = SolpedProceso And PosicionBandeja = PosicionProceso And ProcesoGrupoCompra.Offset(0, 10) = "" Then 'Si el grupo de compra es diferente y estatus es A
                    GoTo SeEncontroSolped2
                End If
            Next ProcesoGrupoCompra
            'LO QUE NO SE ENCONTRO
            RefYaTieneNombre = 0 'Para evitar repetición
                For Each ProcesoGrupoCompra In RangoProceso
                    SolpedProceso = ProcesoGrupoCompra.Offset(0, -7)
                    PosicionProceso = ProcesoGrupoCompra.Offset(0, -6)
                    If ProcesoGrupoCompra.Offset(0, 7) = "A" And SolpedBandeja = SolpedProceso And PosicionBandeja = PosicionProceso And ProcesoGrupoCompra.Offset(0, 10) = "" Then 'Si el grupo de compra es diferente y estatus es A
                        GrupoCompraSuplente = ProcesoGrupoCompra.Value
                        NombreSuplente = ProcesoGrupoCompra.Offset(0, 1)
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 1) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 2) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 3) = SolpedBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 4) = PosicionBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 5) = GrupoCompraSuplente
                        Worksheets("REPORTE").Cells(FilaGrupo, 6) = NombreSuplente
                        RefYaTieneNombre = 1
                    End If

                Next ProcesoGrupoCompra
                    If RefYaTieneNombre = 0 Then
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 1) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 2) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 3) = SolpedBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 4) = PosicionBandeja
                        Worksheets("REPORTE").Cells(FilaGrupo, 5) = "--"
                        Worksheets("REPORTE").Cells(FilaGrupo, 6) = "PETICIÓN ACTIVA" ' La solped esta en estatus A pero no se encuentra en la hoja MM-CO-PA-0004
                    End If
            MensajeTextoBandeja = MensajeTextoBandeja & vbNewLine & GrupoCompra & " " & SolpedBandeja & " " & PosicionBandeja
        End If
        
SeEncontroSolped2:
     
     Next BandejaGrupoCompra
CANCELAR:
'*************************************FIN MODIFICACION PARA MM-CO-PA-0002C (2 PART)****************
    End If
Next Celda
' FIN BUSQUEDA EN BANDEJA

FilaGrupo = 2

Worksheets("REPORTE").Cells(3, 8).Select
Vuelta = 0
For Each CeldaProceso In Rango

Vuelta = Vuelta + 1
Application.StatusBar = "Revisión de Procesos Tratados " & Format((Vuelta / Rango.Count) * 100, "0") & "%"

If IsEmpty(CeldaProceso.Offset(0, -21)) Then
    CeldaT = 0
ElseIf IsNumeric(CeldaProceso.Offset(0, -21)) Then
    CeldaT = CeldaProceso.Offset(0, -21)
Else
    CeldaT = 0
End If

If CeldaT <> CeldaProceso.Value And (CeldaProceso.Offset(0, -22) = "QUIMICOS" Or CeldaProceso.Offset(0, -22) = "PARADAS" Or CeldaProceso.Offset(0, -22) = "RUTINA" Or CeldaProceso.Offset(0, -22) = "CORPORATIVO") Then 'Tienen diferencia en cantidad Tratada
    GrupoCompra = CeldaProceso.Offset(0, -25)
    NombreComprador = CeldaProceso.Offset(0, -24)

    FinalProceso = Worksheets("PET (MM-CO-PA-0004)").Range("I" & Rows.Count).End(xlUp).Row

'**********ESTA SECCIÓN VERIFICA QUE LA SOLPED TRATADA TENGA EL MISMO ESTATUS Y GRUPO DE COMPRA EN LA BANDEJA
    Set RangoProceso = Worksheets("PET (MM-CO-PA-0004)").Range("I2:I" & FinalProceso)
    For Each ProcesoGrupoCompra In RangoProceso 'Grupo de Compra
        If ProcesoGrupoCompra.Value = GrupoCompra And ProcesoGrupoCompra.Offset(0, 7) = "A" And ProcesoGrupoCompra.Offset(0, 10) = "" Then 'Si el grupo de compra es diferente y estatus es A
            SolpedProceso = ProcesoGrupoCompra.Offset(0, -7).Value 'Solped
            PosicionProceso = ProcesoGrupoCompra.Offset(0, -6) 'Posicion
            FinalBandeja = Worksheets("MM-CO-PA-0002C").Range("L" & Rows.Count).End(xlUp).Row
            Set RangoBandeja = Worksheets("MM-CO-PA-0002C").Range("L2:L" & FinalBandeja)
            For Each BandejaGrupoCompra In RangoBandeja
                SolpedBandeja = BandejaGrupoCompra.Offset(0, -9)
                PosicionBandeja = BandejaGrupoCompra.Offset(0, -8)
                If BandejaGrupoCompra.Value = GrupoCompra And BandejaGrupoCompra.Offset(0, 2) = "A" And SolpedProceso = SolpedBandeja And PosicionProceso = PosicionBandeja Then 'Si el grupo de compra es diferente y estatus es A
                    GoTo SeEncontroSolped3
                End If
            Next BandejaGrupoCompra
            '***************************INICIO REVISIÓN BANDEJA 2******************
            FinalBandeja = Worksheets("MM-CO-PA-0002C (2 PART)").Range("L" & Rows.Count).End(xlUp).Row
            If FinalBandeja < 2 Then GoTo SIGUIENTEE
            Set RangoBandeja = Worksheets("MM-CO-PA-0002C (2 PART)").Range("L2:L" & FinalBandeja)
            For Each BandejaGrupoCompra In RangoBandeja
                SolpedBandeja = BandejaGrupoCompra.Offset(0, -9)
                If SolpedBandeja = 1000177293 Then
                    SolpedBandeja = 1000177293
                End If
                PosicionBandeja = BandejaGrupoCompra.Offset(0, -8)
                If BandejaGrupoCompra.Value = GrupoCompra And BandejaGrupoCompra.Offset(0, 2) = "A" And SolpedProceso = SolpedBandeja And PosicionProceso = PosicionBandeja Then 'Si el grupo de compra es diferente y estatus es A
                    GoTo SeEncontroSolped3
                End If
            Next BandejaGrupoCompra
SIGUIENTEE:
            '***************************FIN REVISIÓN BANDEJA 2******************
'**************ESTA SECCIÓN ES PARA LOS PROCESOS TRATADOS PERO CON DIFERENTES GRUPO DE COMPRA QUE EL DE LA BADEJA*****
            'NO SE ENCONTRO EN BANDEJA
            RefYaTieneNombre = 0 'Para evitar repetición
            FinalBandeja = Worksheets("MM-CO-PA-0002C").Range("L" & Rows.Count).End(xlUp).Row
            Set RangoBandeja = Worksheets("MM-CO-PA-0002C").Range("L2:L" & FinalBandeja)
                For Each BandejaGrupoCompra In RangoBandeja
                    SolpedBandeja = BandejaGrupoCompra.Offset(0, -9)
                    PosicionBandeja = BandejaGrupoCompra.Offset(0, -8)
                    If BandejaGrupoCompra.Offset(0, 2) = "A" And SolpedBandeja = SolpedProceso And PosicionBandeja = PosicionProceso Then 'Si el grupo de compra es diferente y estatus es A
                        GrupoCompraSuplente = BandejaGrupoCompra.Value
                        NombreSuplente = BandejaGrupoCompra.Offset(0, 1)
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 8) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 9) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 10) = SolpedProceso
                        Worksheets("REPORTE").Cells(FilaGrupo, 11) = PosicionProceso
                        Worksheets("REPORTE").Cells(FilaGrupo, 12) = GrupoCompraSuplente
                        Worksheets("REPORTE").Cells(FilaGrupo, 13) = NombreSuplente
                        RefYaTieneNombre = 1
                    End If
                Next BandejaGrupoCompra
                
            FinalBandeja = Worksheets("MM-CO-PA-0002C (2 PART)").Range("L" & Rows.Count).End(xlUp).Row
            Set RangoBandeja = Worksheets("MM-CO-PA-0002C (2 PART)").Range("L2:L" & FinalBandeja)
                For Each BandejaGrupoCompra In RangoBandeja
                    SolpedBandeja = BandejaGrupoCompra.Offset(0, -9)
                    PosicionBandeja = BandejaGrupoCompra.Offset(0, -8)
                    If BandejaGrupoCompra.Offset(0, 2) = "A" And SolpedBandeja = SolpedProceso And PosicionBandeja = PosicionProceso Then 'Si el grupo de compra es diferente y estatus es A
                        GrupoCompraSuplente = BandejaGrupoCompra.Value
                        NombreSuplente = BandejaGrupoCompra.Offset(0, 1)
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 8) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 9) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 10) = SolpedProceso
                        Worksheets("REPORTE").Cells(FilaGrupo, 11) = PosicionProceso
                        Worksheets("REPORTE").Cells(FilaGrupo, 12) = GrupoCompraSuplente
                        Worksheets("REPORTE").Cells(FilaGrupo, 13) = NombreSuplente
                        RefYaTieneNombre = 1
                    End If
                Next BandejaGrupoCompra
'*********** ESTA SECCIÓN ES PARA CUANDO EL SOLPED TRATADA NO SE ENCUENTRA EN LA BANDEJA********
                    If RefYaTieneNombre = 0 Then
                        FilaGrupo = FilaGrupo + 1
                        Worksheets("REPORTE").Cells(FilaGrupo, 8) = GrupoCompra
                        Worksheets("REPORTE").Cells(FilaGrupo, 9) = NombreComprador
                        Worksheets("REPORTE").Cells(FilaGrupo, 10) = SolpedProceso
                        Worksheets("REPORTE").Cells(FilaGrupo, 11) = PosicionProceso
                        Worksheets("REPORTE").Cells(FilaGrupo, 12) = "--"
                        Worksheets("REPORTE").Cells(FilaGrupo, 13) = "NO EN BANDEJA"
                    End If
            
            MensajeTextoProceso = MensajeTextoProceso & vbNewLine & GrupoCompra & " " & SolpedProceso & " " & PosicionProceso
        End If
        
SeEncontroSolped3:
     
     Next ProcesoGrupoCompra


' FIN BUSQUEDA EN PROCESO


End If

Next CeldaProceso

MensajeTexto = "NO ESTAN EN PROCESOS TRATADOS" & vbNewLine & MensajeTextoBandeja & vbNewLine & "NO ESTAN EN BANDEJA" & vbNewLine & MensajeTextoProceso
 
'MsgBox MensajeTexto

'VERIFICAR PETICIONES REPETIDAS
Vuelta = 0
Set RangoProceso = Worksheets("PET (MM-CO-PA-0004)").Range("AK2:AK" & FinalProceso)
For Each RefConteo2 In RangoProceso 'RefConteo2

Vuelta = Vuelta + 1
Application.StatusBar = "Revisión Peticiones Repetidas " & Format((Vuelta / RangoProceso.Count) * 100, "0") & "%"

If RefConteo2 = "" Then GoTo Siguiente
Contar = Application.CountIf(Worksheets("PET (MM-CO-PA-0004)").Range("AK2:AK" & FinalProceso), RefConteo2)
If Contar > 1 And RefConteo2.Offset(0, -18) = "" Then

    i = i + 1
    GrupoCompra = RefConteo2.Offset(0, -28)
    NombreComprador = RefConteo2.Offset(0, -27)
    SolpedProceso = RefConteo2.Offset(0, -35)
    PosicionProceso = RefConteo2.Offset(0, -34)

    If i > 1 Then
    For e = 1 To i
        If Repetido(e) = RefConteo2.Value Then GoTo Siguiente
    Next e
    End If
    
    Repetido(i) = RefConteo2.Value
    
    FilaGrupo = FilaGrupo + 1
    Worksheets("REPORTE").Cells(FilaGrupo, 8) = GrupoCompra
    Worksheets("REPORTE").Cells(FilaGrupo, 9) = NombreComprador
    Worksheets("REPORTE").Cells(FilaGrupo, 10) = SolpedProceso
    Worksheets("REPORTE").Cells(FilaGrupo, 11) = PosicionProceso
    Worksheets("REPORTE").Cells(FilaGrupo, 12) = "--"
    Worksheets("REPORTE").Cells(FilaGrupo, 13) = "PETICION REPETIDA"

Siguiente:

End If

Next RefConteo2

Call ProcesosActivos
Call InsertarNuevoGrupo
Call MACROS_3.Solped

Worksheets("REPORTE").Activate
Worksheets("REPORTE").Range("A2:BH65536").Select
If ActiveSheet.AutoFilterMode = True Then
    Selection.AutoFilter
End If
Selection.AutoFilter

Worksheets("Monitoreo").Activate

Application.Calculation = xlCalculationAutomatic

End Sub

Sub GrupoNuevo()

Dim Columna As Range
Dim Celda As Range
Dim Valor As String
Dim FinalBandeja, FinalProcesos As Long
Dim MensajeTextoProceso As Variant
Dim Repetido(1000) As String
Dim i, e As Long


FinalBandeja = Worksheets("MM-CO-PA-0002C").Range("C" & Rows.Count).End(xlUp).Row

Set Columna = Worksheets("MM-CO-PA-0002C").Range("AC2:AC" & FinalBandeja)

For Each Celda In Columna
Valor = Celda
    If Valor = "NO" Then
        'If IsEmpty(i) Then GoTo CONT
        For e = 1 To i
            If Repetido(e) = Celda.Offset(0, -17) Then GoTo CONT
        Next e
        i = 1 + i
        Repetido(i) = Celda.Offset(0, -17)
    MensajeTextoProceso = MensajeTextoProceso & vbNewLine & Celda.Offset(0, -17) & " " & Celda.Offset(0, -16)
CONT:
    End If
Next Celda

'***************INICIO 2 BANDEJA**********
FinalBandeja = Worksheets("MM-CO-PA-0002C (2 PART)").Range("C" & Rows.Count).End(xlUp).Row

If FinalBandeja = 1 Then GoTo CONT5

Set Columna = Worksheets("MM-CO-PA-0002C (2 PART)").Range("AC2:AC" & FinalBandeja)

For Each Celda In Columna
Valor = Celda
    If Valor = "NO" Then
        'If IsEmpty(i) Then GoTo CONT
        For e = 1 To i
            If Repetido(e) = Celda.Offset(0, -17) Then GoTo CONT2
        Next e
        i = 1 + i
        Repetido(i) = Celda.Offset(0, -17)
    MensajeTextoProceso = MensajeTextoProceso & vbNewLine & Celda.Offset(0, -17) & " " & Celda.Offset(0, -16)
CONT2:
    End If
Next Celda
CONT5:
'**************FIN 2 BANDEJA**************

'***************INICIO PROCESOS **********
FinalProcesos = Worksheets("PET (MM-CO-PA-0004)").Range("A" & Rows.Count).End(xlUp).Row

Set Columna = Worksheets("PET (MM-CO-PA-0004)").Range("Y2:Y" & FinalProcesos)

For Each Celda In Columna
Valor = Celda
    If Valor = "NO" Then
        'If IsEmpty(i) Then GoTo CONT
        For e = 1 To i
            If Repetido(e) = Celda.Offset(0, -16) Then GoTo CONT3
        Next e
        i = 1 + i
        Repetido(i) = Celda.Offset(0, -16)
    MensajeTextoProceso = MensajeTextoProceso & vbNewLine & Celda.Offset(0, -16) & " " & Celda.Offset(0, -15)
CONT3:
    End If
Next Celda

'**************FIN PROCESOS **************

If IsEmpty(MensajeTextoProceso) Then GoTo FIN
MsgBox "Los siguientes Grupos de Compras no se encuentran en la lista de Compradores" _
& vbNewLine & MensajeTextoProceso

FIN:
End Sub

Sub VerificarGrupo()
Dim Columna As Range
Dim ColumnaRef As Range
Dim Celda As Range
Dim CeldaRef As Range
Dim Valor As String
Dim FinalBandeja As Long
Dim FinalRef, C As Long
Dim MensajeTextoProceso As String
Dim GrupoCompra As String
Dim i As Long

FinalBandeja = Worksheets("MM-CO-PA-0002C (2)").Range("A" & Rows.Count).End(xlUp).Row
FinalRef = Worksheets("Ref").Range("A" & Rows.Count).End(xlUp).Row
C = FinalRef
Set Columna = Worksheets("MM-CO-PA-0002C (2)").Range("A2:A" & FinalBandeja)
Set ColumnaRef = Worksheets("Ref").Range("A2:A" & FinalRef)

'Borra grupos de compras repetidos en la hoja "Ref"
Worksheets("Ref").Activate

For i = 3 To FinalRef
GrupoCompra = Worksheets("Ref").Cells(i, 1)

If Application.CountIf(Worksheets("Ref").Range(Cells(3, 1), Cells(FinalRef, 1)), GrupoCompra) > 1 Then
    Worksheets("Ref").Cells(i, 1).EntireRow.Delete
End If

Next i

'Borra grupos de compras de procesos anteriores que no fueron guardados para no acumular
For Each CeldaRef In ColumnaRef
    If CeldaRef.Offset(0, 1) = "NO" Then
       CeldaRef.EntireRow.Delete
    End If

Next CeldaRef

'Inserta los Grupo de Compras que no aparecen en la Hoja "Ref"
For Each Celda In Columna
Valor = Celda
    For Each CeldaRef In ColumnaRef
        If Valor = CeldaRef Then GoTo CONT
    Next CeldaRef

C = C + 1
Worksheets("Ref").Range("A" & C) = Celda

CONT:
Next Celda
End Sub

Sub InsertarNuevoGrupo()
Dim FinColumna As Long
Dim FinReporte As Long
Dim ColumnaGrupo, CeldaGrupo As Range

FinColumna = Worksheets("MM-CO-PA-0002C").Range("A" & Rows.Count).End(xlUp).Row

Set ColumnaGrupo = Worksheets("MM-CO-PA-0002C").Range("AC2:AC" & FinColumna)
Worksheets("REPORTE").Cells(3, 15).Select
For Each CeldaGrupo In ColumnaGrupo

    If CeldaGrupo = "NO" Then
        FinReporte = Worksheets("REPORTE").Range("O" & Rows.Count).End(xlUp).Row + 1
        Worksheets("REPORTE").Range("O" & FinReporte) = CeldaGrupo.Offset(0, -17)
        Worksheets("REPORTE").Range("P" & FinReporte) = CeldaGrupo.Offset(0, -16)
        Worksheets("REPORTE").Range("Q" & FinReporte) = CeldaGrupo.Offset(0, -26)
        Worksheets("REPORTE").Range("R" & FinReporte) = CeldaGrupo.Offset(0, -25)
    End If

Next CeldaGrupo

'*************** INICIO 2 BANDEJA****************
FinColumna = Worksheets("MM-CO-PA-0002C (2 PART)").Range("A" & Rows.Count).End(xlUp).Row

Set ColumnaGrupo = Worksheets("MM-CO-PA-0002C (2 PART)").Range("AC2:AC" & FinColumna)
For Each CeldaGrupo In ColumnaGrupo

    If CeldaGrupo = "NO" Then
        FinReporte = Worksheets("REPORTE").Range("O" & Rows.Count).End(xlUp).Row + 1
        Worksheets("REPORTE").Range("O" & FinReporte) = CeldaGrupo.Offset(0, -17)
        Worksheets("REPORTE").Range("P" & FinReporte) = CeldaGrupo.Offset(0, -16)
        Worksheets("REPORTE").Range("Q" & FinReporte) = CeldaGrupo.Offset(0, -26)
        Worksheets("REPORTE").Range("R" & FinReporte) = CeldaGrupo.Offset(0, -25)
    End If

Next CeldaGrupo

'*************** FIN 2 BANDEJA******************

'*************** INICIO 2 BANDEJA****************
FinColumna = Worksheets("PET (MM-CO-PA-0004)").Range("A" & Rows.Count).End(xlUp).Row

Set ColumnaGrupo = Worksheets("PET (MM-CO-PA-0004)").Range("Y2:Y" & FinColumna)
For Each CeldaGrupo In ColumnaGrupo

    If CeldaGrupo = "NO" Then
        FinReporte = Worksheets("REPORTE").Range("O" & Rows.Count).End(xlUp).Row + 1
        Worksheets("REPORTE").Range("O" & FinReporte) = CeldaGrupo.Offset(0, -16)
        Worksheets("REPORTE").Range("P" & FinReporte) = CeldaGrupo.Offset(0, -15)
        Worksheets("REPORTE").Range("Q" & FinReporte) = CeldaGrupo.Offset(0, -23)
        Worksheets("REPORTE").Range("R" & FinReporte) = CeldaGrupo.Offset(0, -22)
    End If

Next CeldaGrupo

'*************** FIN 2 BANDEJA******************


End Sub

Sub ProcesosActivos()

Dim Columna As Range
Dim Celda As Range
Dim FinalProcesos, FinalReporte, A As Long

Application.Calculation = xlCalculationManual

Worksheets("PET (MM-CO-PA-0004)").Activate
FinalProcesos = Worksheets("PET (MM-CO-PA-0004)").Range("A" & Rows.Count).End(xlUp).Row
Set Columna = Worksheets("PET (MM-CO-PA-0004)").Range("P2:P" & FinalProcesos)

Worksheets("REPORTE").Activate
FinalReporte = Worksheets("REPORTE").Range("H" & Rows.Count).End(xlUp).Row + 1

A = 0

For Each Celda In Columna
    If Celda = "A" And Celda.Offset(0, 4) = "X" Then
        Worksheets("REPORTE").Cells(FinalReporte + A, 8) = Celda.Offset(0, -7) 'Grp Compra
        Worksheets("REPORTE").Cells(FinalReporte + A, 9) = Celda.Offset(0, -6) 'Comprador
        Worksheets("REPORTE").Cells(FinalReporte + A, 10) = Celda.Offset(0, -14) * 1 'Solped
        Worksheets("REPORTE").Cells(FinalReporte + A, 11) = Celda.Offset(0, -13) * 1 'Pos
        Worksheets("REPORTE").Cells(FinalReporte + A, 12) = "--" '
        Worksheets("REPORTE").Cells(FinalReporte + A, 13) = "Status A, Solped Borrada" '
        A = A + 1
    End If
Next Celda

For Each Celda In Columna
    If Celda = "A" And Celda.Offset(0, 1) = "X" Then
        Worksheets("REPORTE").Cells(FinalReporte + A, 8) = Celda.Offset(0, -7) 'Grp Compra
        Worksheets("REPORTE").Cells(FinalReporte + A, 9) = Celda.Offset(0, -6) 'Comprador
        Worksheets("REPORTE").Cells(FinalReporte + A, 10) = Celda.Offset(0, -14) * 1 'Solped
        Worksheets("REPORTE").Cells(FinalReporte + A, 11) = Celda.Offset(0, -13) * 1 'Pos
        Worksheets("REPORTE").Cells(FinalReporte + A, 12) = "--" '
        Worksheets("REPORTE").Cells(FinalReporte + A, 13) = "Status A, Solped Concluida" 'Comprador
        A = A + 1
    End If
Next Celda

End Sub


