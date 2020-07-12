Attribute VB_Name = "MACROS"

Sub BackLog_Mensual()
Attribute BackLog_Mensual.VB_ProcData.VB_Invoke_Func = " \n14"
Dim FilaFinal As Long
On Error Resume Next
Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Worksheets("MM-CO-PA-0002C").Activate
    
    If ThisWorkbook.Worksheets("MM-CO-PA-0002C").FilterMode = True Then
        Worksheets("MM-CO-PA-0002C").ShowAllData
    End If
    
    'CONTEO ULTIMA FILA
    FilaFinal = Worksheets("MM-CO-PA-0002C").Range("C" & Rows.Count).End(xlUp).Row
    
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    Sheets("MM-CO-PA-0002C").Select
    Columns("A:B").Select
    Selection.ClearContents
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Conteo"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = 1
    Selection.Copy
    Range("A2:A" & FilaFinal).Select
    ActiveSheet.Paste
    
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Md"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=+IF(MID(RC[15],1,2)=""RB"",511,516)"
    Selection.Copy
    Range("B2:B" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Rango Antiguedad
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "Rg. Ant"
    Range("AA2").Select
    ActiveCell.FormulaLocal = "=+SI(R2<=30;""<= 30 días"";SI(R2>30;SI(R2<=60;""31 a 60 días"";SI(R2>60;SI(R2<=90;""61 a 90 días"";""> a 90 días"")))))"
    Selection.Copy
    Range("AA2:AA" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Equipo
    Range("AB1").Select
    ActiveCell.FormulaR1C1 = "Equipo"
        
    'Equipo
    Range("AB1").Select
    ActiveCell.FormulaR1C1 = "Equipo"
    Range("AB2").Select
    ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV($L2;'H:\INFORME GESTION\REISEL SANCHEZ\07 DATA\[Compradores por Equip Procura.xls]Compradores'!$A:$E;4;FALSO));""NO"";BUSCARV($L2;'H:\INFORME GESTION\REISEL SANCHEZ\07 DATA\[Compradores por Equip Procura.xls]Compradores'!$A:$E;4;FALSO))"
    Selection.Copy
    Range("AB2:AB" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'Superint
    Range("AC1").Select
    ActiveCell.FormulaR1C1 = "Superint"
    Range("AC2").Select
    ActiveCell.FormulaLocal = "=+SI(ESERROR(BUSCARV($L2;'[Compradores por Equip Procura.xls]Compradores'!$A:$E;5;FALSO));""NO"";BUSCARV($L2;'[Compradores por Equip Procura.xls]Compradores'!$A:$E;5;FALSO))"
    Selection.Copy
    Range("AC2:AC" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Periodo
    Range("AD1").Select
    ActiveCell.FormulaR1C1 = "Periodo"
    Range("AD2").Select
    ActiveCell.FormulaLocal = "=+AÑO(I2)&"" ""&MES(I2)"
    Selection.Copy
    Range("AD2:AD" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Referencias
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "Referencias"
    Range("AE2").Select
    ActiveCell.FormulaLocal = "=AB2&""-""&N2"
    Selection.Copy
    Range("AE2:AE" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
       
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "Referencias 2"
    Range("AF2").Select
    ActiveCell.FormulaLocal = "=L2&""-""&N2"
    Selection.Copy
    Range("AF2:AF" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "Ref Cuadros"
    Range("AG2").Select
    ActiveCell.FormulaLocal = "=AC2&""-""&AA2"
    Selection.Copy
    Range("AG2:AG" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "Ref Cuadro 2"
    Range("AH2").Select
    ActiveCell.FormulaLocal = "=AG2&""-""&N2"
    Selection.Copy
    Range("AH2:AH" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        

    'CALCULAR
    Worksheets("MM-CO-PA-0002C").Calculate
    Application.Calculation = xlCalculationAutomatic
    
    'PEGAR VALORES
     Range("B2:B" & FilaFinal).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    

    Range("AA2:AH" & FilaFinal).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    
    Worksheets("MM-CO-PA-0002C").Select
    Range("A1").Select

    
'************************** 2 PART*******************************
    Worksheets("MM-CO-PA-0002C (2 PART)").Activate
    
    If ThisWorkbook.Worksheets("MM-CO-PA-0002C (2 PART)").FilterMode = True Then
        Worksheets("MM-CO-PA-0002C (2 PART)").ShowAllData
    End If
    
    'CONTEO ULTIMA FILA
    FilaFinal = Worksheets("MM-CO-PA-0002C (2 PART)").Range("C" & Rows.Count).End(xlUp).Row
    
    If FilaFinal = 1 Then GoTo CONT4
    
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    Sheets("MM-CO-PA-0002C (2 PART)").Select
    Columns("A:B").Select
    Selection.ClearContents
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Conteo"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = 1
    Selection.Copy
    Range("A2:A" & FilaFinal).Select
    ActiveSheet.Paste
    
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Md"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=+IF(MID(RC[15],1,2)=""RB"",511,516)"
    Selection.Copy
    Range("B2:B" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Rango Antiguedad
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "Rg. Ant"
    Range("AA2").Select
    ActiveCell.FormulaLocal = "=+SI(R2<=30;""<= 30 días"";SI(R2>30;SI(R2<=60;""31 a 60 días"";SI(R2>60;SI(R2<=90;""61 a 90 días"";""> a 90 días"")))))"
    Selection.Copy
    Range("AA2:AA" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Equipo
    Range("AB1").Select
    ActiveCell.FormulaR1C1 = "Equipo"
        
    'Equipo
    Range("AB1").Select
    ActiveCell.FormulaR1C1 = "Equipo"
    Range("AB2").Select
    ActiveCell.FormulaLocal = "=+BUSCARV($L2;'[Compradores por Equip Procura.xls]Compradores'!$A:$E;4;FALSO)"
    Selection.Copy
    Range("AB2:AB" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'Superint
    Range("AC1").Select
    ActiveCell.FormulaR1C1 = "Superint"
    Range("AC2").Select
    ActiveCell.FormulaLocal = "=+SI(ESERROR(BUSCARV($L2;'[Compradores por Equip Procura.xls]Compradores'!$A:$E;5;FALSO));""NO"";BUSCARV($L2;'[Compradores por Equip Procura.xls]Compradores'!$A:$E;5;FALSO))"
    Selection.Copy
    Range("AC2:AC" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Periodo
    Range("AD1").Select
    ActiveCell.FormulaR1C1 = "Periodo"
    Range("AD2").Select
    ActiveCell.FormulaLocal = "=+AÑO(I2)&"" ""&MES(I2)"
    Selection.Copy
    Range("AD2:AD" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Referencias
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "Referencias"
    Range("AE2").Select
    ActiveCell.FormulaLocal = "=AB2&""-""&N2"
    Selection.Copy
    Range("AE2:AE" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
       
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "Referencias 2"
    Range("AF2").Select
    ActiveCell.FormulaLocal = "=L2&""-""&N2"
    Selection.Copy
    Range("AF2:AF" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "Ref Cuadros"
    Range("AG2").Select
    ActiveCell.FormulaLocal = "=AC2&""-""&AA2"
    Selection.Copy
    Range("AG2:AG" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "Ref Cuadro 2"
    Range("AH2").Select
    ActiveCell.FormulaLocal = "=AG2&""-""&N2"
    Selection.Copy
    Range("AH2:AH" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False


    'CALCULAR
    Worksheets("MM-CO-PA-0002C (2 PART)").Calculate
    Application.Calculation = xlCalculationAutomatic
    
    'PEGAR VALORES
     Range("B2:B" & FilaFinal).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    

    Range("AA2:AH" & FilaFinal).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    
    Worksheets("MM-CO-PA-0002C (2 PART)").Select
    Range("A1").Select


CONT4:
'************************** 2 PART*******************************
    
    ThisWorkbook.Activate
    
    'ACTUALIZAR TABLAS DINAMICAS
    Sheets("BLACKLOG").Select
    Range("A7").Select
    ActiveSheet.PivotTables("Tabla dinámica1").PivotCache.Refresh
    Sheets("Status N").Select
    Range("A7").Select
    ActiveSheet.PivotTables("Tabla dinámica2").PivotCache.Refresh
    Sheets("Status A").Select
    Range("B7").Select
    ActiveSheet.PivotTables("Tabla dinámica2").PivotCache.Refresh
    


End Sub

Sub PeticionesOfertas()
On Error Resume Next
Dim FilaFinal As Long
Application.ScreenUpdating = False
    
    Application.Calculation = xlCalculationManual
    
    Worksheets("PET (MM-CO-PA-0004)").Activate
    
    If ThisWorkbook.Worksheets("PET (MM-CO-PA-0004)").FilterMode = True Then
        Worksheets("PET (MM-CO-PA-0004)").ShowAllData
    End If
    
    'CONTEO ULTIMA FILA
    FilaFinal = Worksheets("PET (MM-CO-PA-0004)").Range("A" & Rows.Count).End(xlUp).Row
    
    'Modalidad
    Range("X1").Select
    ActiveCell.FormulaR1C1 = "Modalidad"
    Range("X2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=+SI(H2=""ZANA"";""MENOR"";SI(H2=""ZANC"";""MAYOR"";""EXTERIOR""))"
    Selection.Copy
    Range("X2:X" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Equipo
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "Equipo"
    Range("Y2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(I2;'H:\INFORME GESTION\REISEL SANCHEZ\07 DATA\[Compradores por Equip Procura.xls]Compradores'!$A:$D;4;FALSO));""NO"";BUSCARV(I2;'H:\INFORME GESTION\REISEL SANCHEZ\07 DATA\[Compradores por Equip Procura.xls]Compradores'!$A:$D;4;FALSO))"
    Selection.Copy
    Range("Y2:Y" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'MES
    Range("Z1").Select
    ActiveCell.FormulaR1C1 = "MES"
    Range("Z2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=MES(G2)"
    Selection.Copy
    Range("Z2:Z" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'AÑO
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "AÑO"
    Range("AA2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=AÑO(G2)"
    Selection.Copy
    Range("AA2:AA" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'REF
    Range("AB1").Select
    ActiveCell.FormulaR1C1 = "REF"
    Range("AB2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=X2&""-""&I2&""-""&AA2"
    Selection.Copy
    Range("AB2:AB" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'REF 2
    Range("AC1").Select
    ActiveCell.FormulaR1C1 = "REF 2"
    Range("AC2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=AB2&""-""&Z2"
    Selection.Copy
    Range("AC2:AC" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'REF 3
    Range("AD1").Select
    ActiveCell.FormulaR1C1 = "REF 3"
    Range("AD2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=X2&""-""&Y2&""-""&AA2&""-""&Z2"
    Selection.Copy
    Range("AD2:AD" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'REF 4
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "REF 4"
    Range("AE2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=X2&""-""&Y2&""-""&AA2"
    Selection.Copy
    Range("AE2:AE" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Peticiónes Activas
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "Peticiónes Activas"
    Range("AF2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(Y(CONTAR.SI($D$2:D2;D2)=1;Q2="""";S2="""";T2="""");1;0)"
    Selection.Copy
    Range("AF2:AF" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Peticiónes Activas
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "Peticiónes Activas"
    Range("AF2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(Y(CONTAR.SI($AJ$2:AJ2;AJ2)=1;P2=""A"";Q2="""";S2="""";T2="""");1;0)"
    Selection.Copy
    Range("AF2:AF" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Peticiones Realizadas
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "Peticiones Realizadas"
    Range("AG2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(CONTAR.SI($D$2:D2;D2)=1;1;0)"
    Selection.Copy
    Range("AG2:AG" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Posiciones Activas
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "Posiciones Activas"
    Range("AH2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(AF2=1;CONTAR.SI(AJ:AJ;AJ2);0)"
    Selection.Copy
    Range("AH2:AH" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Posiciones Activas 2
    Range("AI1").Select
    ActiveCell.FormulaR1C1 = "Posiciones Activas 2"
    Range("AI2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(Y(P2=""A"";Q2="""";S2="""";T2="""");1;0)"
    Selection.Copy
    Range("AI2:AI" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Ref Conteo
    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = "Ref Conteo"
    Range("AJ2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(Y(P2=""A"";T2=""X"");D2&""-X"";D2&""-""&P2)"
    Selection.Copy
    Range("AJ2:AJ" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Ref Conteo Solped Repetidas
    Range("AK1").Select
    ActiveCell.FormulaR1C1 = "Ref Conteo 2"
    Range("AK2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(O(S2=""L"";P2=""B"";P2=""N"";T2=""X"");"""";B2&""-""&C2)"
    Selection.Copy
    Range("AK2:AK" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'Ref Conteo Solped Repetidas
    Range("AL1").Select
    ActiveCell.FormulaR1C1 = "Ref Grupo"
    Range("AL2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=X2&""-""&I2"
    Selection.Copy
    Range("AL2:AL" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False


    'CALCULAR
    Worksheets("PET (MM-CO-PA-0004)").Calculate
    Application.Calculation = xlCalculationAutomatic
    
    'PEGAR VALORES
     Range("X2:AL" & FilaFinal).Select
    Selection.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    

    'AJUSTAR TABLAS DINAMICAS
    Worksheets("TABLAS").PivotTables("Tabla dinámica3").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica3").PivotFields("Petición 1"), "Cuenta de Petición 1", xlCount

    Worksheets("TABLAS").PivotTables("Tabla dinámica4").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica4").PivotFields("Petición 1"), "Cuenta de Petición 1", xlCount
    
    Worksheets("TABLAS").PivotTables("Tabla dinámica4").PivotFields("Cuenta de Petición 1"). _
        Function = xlSum
        
    Worksheets("TABLAS").PivotTables("Tabla dinámica3").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica4").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica5").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica6").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica7").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica8").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica9").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica10").PivotCache.Refresh
        
    
    Worksheets("PET (MM-CO-PA-0004)").Range("A1").Select

End Sub


Sub Ref()
On Error Resume Next
Dim FilaFinal As Long
Application.ScreenUpdating = False
    
    Application.Calculation = xlCalculationManual
    
    Worksheets("Ref").Activate
    
    If ThisWorkbook.Worksheets("Ref").FilterMode = True Then
        Worksheets("Ref").ShowAllData
    End If
    
    'CONTEO ULTIMA FILA
    FilaFinal = Worksheets("Ref").Range("A" & Rows.Count).End(xlUp).Row
        
    'EQUIPO
    Range("B3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(A3;'[Compradores por Equip Procura.xls]Compradores'!$A:$D;4;FALSO));""NO"";BUSCARV(A3;'[Compradores por Equip Procura.xls]Compradores'!$A:$D;4;FALSO))"
    Selection.Copy
    Range("B3:B" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'EQUIPO
    Range("C3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=BUSCARV(A3;'[Compradores por Equip Procura.xls]Compradores'!$A:$E;5;FALSO)"
    Selection.Copy
    Range("C3:C" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'A
    Range("D3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=CONTAR.SI('MM-CO-PA-0002C'!AF:AF;$A3&""-A"")+CONTAR.SI('MM-CO-PA-0002C (2 PART)'!AF:AF;$A3&""-A"")"
    Selection.Copy
    Range("D3:D" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'N
    Range("E3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=CONTAR.SI('MM-CO-PA-0002C'!AF:AF;$A3&""-N"")+CONTAR.SI('MM-CO-PA-0002C (2 PART)'!AF:AF;$A3&""-N"")"
    Selection.Copy
    Range("E3:E" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'TOTAL
    Range("F3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=D3+E3"
    Selection.Copy
    Range("F3:F" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False


    'MAYOR-POSICIONES
    Range("G3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(A3;TABLAS!A:B;2;FALSO));0;BUSCARV(A3;TABLAS!A:B;2;FALSO))"
    Selection.Copy
    Range("G3:G" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'MENOR-POSICIONES
    Range("H3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(A3;TABLAS!A:C;3;FALSO));0;BUSCARV(A3;TABLAS!A:C;3;FALSO))"
    Selection.Copy
    Range("H3:H" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'MAYOR-PETICIONES
    Range("I3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(A3;TABLAS!H:I;2;FALSO));0;BUSCARV(A3;TABLAS!H:I;2;FALSO))"
    Selection.Copy
    Range("I3:I" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'MENOR-PETICIONES
    Range("J3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(A3;TABLAS!H:J;3;FALSO));0;BUSCARV(A3;TABLAS!H:J;3;FALSO))"
    Selection.Copy
    Range("J3:J" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'PEDIDO
    Range("K3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(A3;TABLAS!AA:AB;2;FALSO));0;BUSCARV(A3;TABLAS!AA:AB;2;FALSO))"
    Selection.Copy
    Range("K3:K" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    'POS PEDIDO
    Range("L3").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(A3;TABLAS!X:Y;2;FALSO));0;BUSCARV(A3;TABLAS!X:Y;2;FALSO))"
    Selection.Copy
    Range("L3:L" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False


    'CALCULAR
    Worksheets("MM-CO-PO-0019").Calculate
    Application.Calculation = xlCalculationAutomatic
    
    'PEGAR VALORES
     Range("A3:L" & FilaFinal).Select
    Selection.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    
    Worksheets("Ref").Range("A1").Select
    
    ThisWorkbook.Activate

End Sub

Sub Actualizar()
Application.ScreenUpdating = False
Dim Answer As Integer
Dim TotalMonitoreo, TotalBacklog As Long

On Error Resume Next
Application.StatusBar = "Procesando..."
Answer = MsgBox("¿Desea actualizar todos los datos?", vbYesNo)

If Answer = vbNo Then GoTo CANCELAR

Worksheets("Monitoreo").Range("B1") = Date

'ABRIR LIBRO DE COMPRADORES *** INICIO***
For Each Libro In Workbooks
    If Libro.Name = "Compradores por Equip Procura.xls" Then
        GoTo AAA
    End If
Next Libro
    
    If CreateObject("WScript.Network").UserName = "REISEL SANCHEZ" Then
    Workbooks.Open Filename:= _
    "D:\Desktop\00 GESTION\07 DATA\Compradores por Equip Procura.xls", Password:="BARIVEN"
        GoTo BBB
    End If
    
    Workbooks.Open Filename:= _
    "H:\INFORME GESTION\REISEL SANCHEZ\07 DATA\Compradores por Equip Procura.xls", Password:="BARIVEN"

BBB:
For Each Libro In Workbooks
    If Libro.Name = "Compradores por Equip Procura.xls" Then
        GoTo AAA
    End If
Next Libro
        MsgBox ("NO SE ENCUENTRA ARCHIVO ""..Compradores por Equipo Procura.."" EN INFORME DE GESTIÓN (PUBLICO)" & vbNewLine & vbNewLine & _
        "PARA LA EJECUCIÓN DE LA MACRO SE REQUIERE TENER ABIERTO EL ARCHIVO ""..Compradores por Equipo Procura..""")
        GoTo CANCELAR
AAA:

    ThisWorkbook.Activate
    
'ABRIR LIBRO DE COMPRADORES *** FINAL***

Call BackLog_Mensual
Call PeticionesOfertas
Call Pedido31
Call MACROS_2.VerificarGrupo
Call Ref

For Each Libro In Workbooks
    If Libro.Name = "Compradores por Equip Procura.xls" Then
        Workbooks("Compradores por Equip Procura.xls").Close
    End If
Next Libro

Worksheets("Monitoreo").Select
'*********VERIFICACIÓN DE DATOS BLACKLOG******************
Dim UltimaFilaM, UltimaFilaB As Long
Dim RANGOM As Range
Dim CELDAM As Range
Dim REFM, REFB As Long

UltimaFilaM = Worksheets("Monitoreo").Range("G" & Rows.Count).End(xlUp).Row
Set RANGOM = Worksheets("Monitoreo").Range("G2:G" & UltimaFilaM)
For Each CELDAM In RANGOM
    If CELDAM = "Backlog Total" Then
        REFM = CELDAM.Row + 1
    End If
Next CELDAM
UltimaFilaB = Worksheets("BLACKLOG").Range("E" & Rows.Count).End(xlUp).Row
Set RANGOM = Worksheets("BLACKLOG").Range("E2:E" & UltimaFilaB)
For Each CELDAM In RANGOM
    If CELDAM = "Backlog Total" Then
        REFB = CELDAM.Row
    End If
Next CELDAM

TotalMonitoreo = Worksheets("Monitoreo").Cells(REFM, 7).Value
TotalBacklog = Worksheets("BLACKLOG").Cells(REFB, 4).Value


If Not TotalMonitoreo = TotalBacklog Then
    MsgBox ("LOS NUMEROS DE BLACKLOG TOTALES NO COINCIDEN, SE DEBE VERIFICAR QUE TODOS LOS GRUPOS DE COMPRAS OBTENIDOS EN EL QUERY ESTEN REGISTRADOS EN LA HOJA Ref Y EN EL ARCHIVO Compradores por Equip Procura UBICADO EN H:\INFORME GESTION\REISEL SANCHEZ\07 DATA")
End If

Call MACROS_2.GrupoNuevo
'*********VERIFICACIÓN DE DATOS BLACKLOG******************
CANCELAR:

Application.StatusBar = False

End Sub

Sub pRUEBA()

Dim FilaFinal As Long
Dim Rango As Range
Dim Celda As Range

FilaFinal = Worksheets("Ref").Range("A" & Rows.Count).End(xlUp).Row

Set Rango = Application.Workbooks("Ref").Range("A3:A" & FilaFinal)

For Each Celda In Rango.Cells
    If IsError(WorksheetFunction.VLookup(Celda, Workbooks("Compradores por Equip Procura.xls").Workbooks("Compradores").Range("A:A"), 1, False)) Then
        
        
        If IsError(WorksheetFunction.VLookup(Celda, Workbooks("Compradores por Equip Procura.xls").Workbooks("Compradores").Range("A:A"), 1, False)) Then

Next Celda

End Sub

Sub Limpiar()
On Error Resume Next
Dim Answer As Integer
Dim T, Z, X As Long


Answer = MsgBox("¿Desea BORRAR todos los datos?", vbYesNo)

If Answer = vbNo Then GoTo CANCELAR

    Sheets("MM-CO-PA-0002C").Select
    ActiveSheet.ShowAllData
    Cells.Select
    Range("A1").Activate
    Selection.ClearContents
    Range("A2").Select
    Cells.Select
    
    Sheets("MM-CO-PA-0002C (2 PART)").Select
    ActiveSheet.ShowAllData
    Cells.Select
    Range("A1").Activate
    Selection.ClearContents
    Range("A2").Select
    Cells.Select
    
    Sheets("MM-CO-PA-0002C (2)").Select
    ActiveSheet.ShowAllData
    Cells.Select
    Selection.ClearContents
    Range("A2").Select
    Cells.Select
    
    Sheets("PET (MM-CO-PA-0004)").Select
    ActiveSheet.ShowAllData
    Cells.Select
    Range("A1").Activate
    Selection.ClearContents
    Range("A2").Select
    Cells.Select
    
    Sheets("MM-CO-PO-0031").Select
    ActiveSheet.ShowAllData
    Cells.Select
    Range("A1").Activate
    Selection.ClearContents
    Range("A2").Select
    Cells.Select
    
    Sheets("PDC").Select
    ActiveSheet.ShowAllData
    Cells.Select
    Range("A1").Activate
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 0
    End With
    ActiveWindow.FreezePanes = False
    Range("A2").Select
    Cells.Select
    Sheets("Monitoreo").Select
      
    
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

CANCELAR:
End Sub

Sub Pedido31()

On Error Resume Next
Dim FilaFinal As Long
Application.ScreenUpdating = False
    
    Application.Calculation = xlCalculationManual
    
    Worksheets("MM-CO-PO-0031").Activate
    
    If ThisWorkbook.Worksheets("MM-CO-PO-0031").FilterMode = True Then
        Worksheets("MM-CO-PO-0031").ShowAllData
    End If
    
    'CONTEO ULTIMA FILA
    FilaFinal = Worksheets("MM-CO-PO-0031").Range("A" & Rows.Count).End(xlUp).Row
    
    
    'Año
    Range("Z2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=+AÑO(F2)"
    Selection.Copy
    Range("Z2:Z" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Mes
    Range("AA2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=+MES(F2)"
    Selection.Copy
    Range("AA2:AA" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        
    'Pedido
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "Pedido"
    Range("AE2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=+SI(AF2=1;1;0)"
    Selection.Copy
    Range("AE2:AE" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Ref
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "Ref"
    Range("AF2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=+CONTAR.SI($G$2:G2;G2)"
    Selection.Copy
    Range("AF2:AF" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Com x Mes
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "Com x Mes"
    Range("AG2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=T2&""-""&AÑO(F2)&""-""&MES(F2)"
    Selection.Copy
    Range("AG2:AG" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Grp x Mes
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "Grp x Mes"
    Range("AH2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=BUSCARV(T2;Ref!A:B;2;FALSO)&""-""&AÑO(F2)&""-""&MES(F2)"
    Selection.Copy
    Range("AH2:AH" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Pedido por Año
    Range("AI1").Select
    ActiveCell.FormulaR1C1 = "Pedido por Año"
    Range("AI2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=T2&""-""&AÑO(F2)"
    Selection.Copy
    Range("AI2:AI" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'Grp x Año
    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = "Grp x Año"
    Range("AJ2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=BUSCARV(T2;Ref!A:B;2;FALSO)&""-""&AÑO(F2)"
    Selection.Copy
    Range("AJ2:AJ" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'NOMBRE
    Range("AK1").Select
    ActiveCell.FormulaR1C1 = "NOMBRE"
    Range("AK2").Select
    Selection.NumberFormat = "0"
    ActiveCell.FormulaLocal = "=SI(ESERROR(BUSCARV(A2*1;Empresas!A:B;2;FALSO));B2;BUSCARV(A2*1;Empresas!A:B;2;FALSO))"
    Selection.Copy
    Range("AK2:AK" & FilaFinal).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'CALCULAR
    Worksheets("MM-CO-PO-0031").Calculate
    Application.Calculation = xlCalculationAutomatic
    
    'PEGAR VALORES
     Range("Z2:AK" & FilaFinal).Select
    Selection.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    
    
    Worksheets("TABLAS").PivotTables("Tabla dinámica3").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica4").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica5").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica6").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica7").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica8").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica9").PivotCache.Refresh
    Worksheets("TABLAS").PivotTables("Tabla dinámica1").PivotCache.Refresh
    
    Worksheets("MM-CO-PO-0031").Range("A1").Select

End Sub
