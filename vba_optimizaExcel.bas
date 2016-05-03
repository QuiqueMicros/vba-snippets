Attribute VB_Name = "modOptimize"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           >>> binarytopic.com <<<                            '
'                            coded by Diego F.C.                               '
'     http://binarytopic.com/optimizar-velocidad-de-calculo-de-excel-vba/      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Sub optimizaExcel(Activar As Boolean, Optional Agresivo As Boolean = False)
'Optimiza Excel para agilizar la velocidad de cálculo y evitar bloqueos de por
' refresco de pantalla, consumo de memoria...
'
' >>>> binarytopic.com <<<<<
'
' ARGUMENTOS:
'   Activar: Boolean. Activa o desactiva los parámetros a optimizar.
'   Aggresivo: Boolean. Habilita la otimización de forma agresiva. Deshabilita
'       cache de tablas dinámicas, guardado de datos...
    
    Dim WS As Worksheet
    Dim PVT As PivotTable

    If Activar Then
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        Application.CutCopyMode = False
        If Agresivo Then
            For Each WS In ThisWorkbook.Worksheets
                If ThisWorkbook.PivotCaches.Count > 1 Then
                    For Each PVT In WS.PivotTables
                        PVT.CacheIndex = 1
                        PVT.SaveData = False
                    Next PVT
                End If
            Next WS
        End If
    Else
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.CutCopyMode = True
    End If
End Sub
