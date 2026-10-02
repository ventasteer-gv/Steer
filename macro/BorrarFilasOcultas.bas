Attribute VB_Name = "ModCatalogo"
Option Explicit

' Borra de la tabla de la hoja "Productos" todas las filas que estan
' OCULTAS por el filtro (o por las segmentaciones) y deja solo las visibles.
Sub BorrarFilasOcultas()
    Dim ws As Worksheet, lo As ListObject, sc As SlicerCache
    Dim n As Long, i As Long, j As Long
    Dim nBorrar As Long, nVisibles As Long
    Dim borrar() As Boolean
    Dim calcAnterior As XlCalculation

    Set ws = ThisWorkbook.Worksheets("Productos")
    Set lo = ws.ListObjects(1)   ' tabla Catálogo_Python_HTML

    If lo.DataBodyRange Is Nothing Then
        MsgBox "La tabla está vacía.", vbInformation
        Exit Sub
    End If

    ' 1) Anotar qué filas están ocultas
    n = lo.ListRows.Count
    ReDim borrar(1 To n)
    For i = 1 To n
        If lo.DataBodyRange.Rows(i).EntireRow.Hidden Then
            borrar(i) = True
            nBorrar = nBorrar + 1
        End If
    Next i
    nVisibles = n - nBorrar

    If nBorrar = 0 Then
        MsgBox "No hay filas ocultas. Primero aplica el filtro.", vbInformation
        Exit Sub
    End If
    If nVisibles = 0 Then
        MsgBox "El filtro no deja ninguna fila visible. No se borra nada.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Se CONSERVARÁN " & nVisibles & " filas visibles." & vbCrLf & _
              "Se BORRARÁN " & nBorrar & " filas ocultas." & vbCrLf & vbCrLf & _
              "¿Continuar?", vbYesNo + vbExclamation, "Borrar filas ocultas") <> vbYes Then Exit Sub

    calcAnterior = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 2) Quitar filtros y segmentaciones (para borrar en bloques sin riesgo)
    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo Fin

    ' 3) Borrar de abajo hacia arriba, por bloques contiguos (rápido)
    i = n
    Do While i >= 1
        If borrar(i) Then
            j = i
            Do While j > 1
                If Not borrar(j - 1) Then Exit Do
                j = j - 1
            Loop
            ws.Range(lo.DataBodyRange.Rows(j), lo.DataBodyRange.Rows(i)).Delete Shift:=xlShiftUp
            i = j - 1
        Else
            i = i - 1
        End If
    Loop

Fin:
    Application.Calculation = calcAnterior
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbCritical
    Else
        MsgBox "Listo. Quedaron " & lo.ListRows.Count & " filas en el catálogo.", vbInformation
    End If
End Sub
