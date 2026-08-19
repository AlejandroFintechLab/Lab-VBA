Attribute VB_Name = "Módulo1"

'1.CONDICIONALES If, ElseIf y Else

Public Sub If_Statements()
    ' Definiendo mis variables
    celda = ActiveCell.Value
    ' Verificando las condiciones
    If celda >= 70 Then
        MsgBox ("Eres mayor de 70.")
    ElseIf celda < 70 And celda >= 21 Then
        MsgBox ("Eres mayor de 21.")
    Else
        MsgBox ("Eres menor de edad.")
    End If
    
End Sub

'2. CONDICIONALES Select Case Is y Else

Public Sub Select_Case_Statements()
    'Definiendo los casos
    Select Case ActiveCell.Value
        Case Is > 70
            MsgBox ("Mayor de 70")
        Case 21 To 69
            MsgBox ("Mayor de 21")
        Case Else
            MsgBox ("Menor de edad")
    End Select
            
End Sub

'3. CICLOS Do While

Public Sub Do_While_Loops()
    'Declarando el contador o acumulador
    'i = 0
    
    'Estableciendo el ciclo}
    Do While ActiveCell.Value <> "" 'i <= 10
        'Activando las sentencias condicionales
        If_Statements
        'Moviéndome por la columna.
        ActiveCell.Offset(1, 0).Select
        'i = i + 1
    Loop
End Sub


'4. CICLOS For Each

Public Sub For_Each_Loops()
    'Definiendo mis variables
    Dim user As Range
    
    'Declarando el bucle para cada elemento
    For Each user In Selection
        'Activando las sentencias condicionales
        If_Statements
        'Moviéndome por la columna.
        ActiveCell.Offset(1, 0).Select
    Next user
    
End Sub


'5. CICLOS For Next

Public Sub For_Next_Loop()
    ' 1. Declarando mis variables
    Dim i As Integer
    
    '2. Creando el bucle For Next
    For i = 1 To ActiveSheet.UsedRange.Rows.Count - 1
        'Activando las sentencias condicionales
        If_Statements
        'Moviéndome por la columna.
        ActiveCell.Offset(1, 0).Select
    Next i
    


End Sub
