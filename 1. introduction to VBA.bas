Attribute VB_Name = "Módulo1"

'1.CONDICIONALES If, ElseIf y Else
' Subproceso que compara los valores de la celda activa y muestra un mensaje en
' consecuencia

Public Sub If_Statements()

    ' 1. Definiendo mis variables: creo la variable celda para acceder al valor (.Value)
    ' de la celda activa (ActiveCell)

    celda = ActiveCell.Value

    ' 2. Verificando las condiciones:
    ' - valor de la celda >= 70 --> Muestra "Eres mayor de 70."
    ' - 60 > valor de la celda >= 21 --> Muestra "Eres mayor de 21."
    ' - De otro modo (cualquier otro caso) --> Muestra "Eres menor de edad."

    If celda >= 60 Then
        MsgBox ("Eres mayor de 60.")
    ElseIf celda < 60 And celda >= 21 Then
        MsgBox ("Eres mayor de 21.")
    Else
        MsgBox ("Eres menor de edad.")
    End If
    
End Sub

'2. CONDICIONALES Select Case Is y Else
' Ejercicio similar al #1, que emplea las sentencias Select Case en lugar de 
' If/ElseIf/Else

Public Sub Select_Case_Statements()

    'Definiendo los casos: Establesco los casos para el valor 
    ' de la variable (ActiveCell.Value) y que sucede en cada uno.

    ' - valor de la celda >= 70 --> Muestra "Eres mayor de 70."
    ' - 60 > valor de la celda >= 21 --> Muestra "Eres mayor de 21."
    ' - De otro modo (cualquier otro caso) --> Muestra "Eres menor de edad."


    Select Case ActiveCell.Value
        Case Is > 60
            MsgBox ("Mayor de 60")
        Case 21 To 59
            MsgBox ("Mayor de 21")
        Case Else
            MsgBox ("Menor de edad")
    End Select
            
End Sub

'3. CICLOS Do While-Usando un Contador
' Ciclo que activa el subproceso "If_Statements" un número específico de veces,
' definidas por la variable contador.

Public Sub Do_While_Loops_a()
    'Declarando el contador o acumulador
    i = 0
    
    'Estableciendo el ciclo: Repite el código interno mientras el contador sea menor e igual
    ' a 10, 

    Do While i <= 10
        'Activando las sentencias condicionales
        If_Statements

        'Moviéndome por la columna: El método Offset nos permite movernos por las filas
        ' y columnas (columnas, filas).
        ActiveCell.Offset(1, 0).Select
        
        'Incrementamos en 1 el contador, para evitar que el ciclo se extienda para siempre.
        i = i + 1
    Loop
End Sub


'4. CICLOS Do While-Versión sin Contador
' Ciclo similar al anterior, que cambia la lógica del código para no usar un contador.

Public Sub Do_While_Loops_b()

    'Estableciendo el ciclo
    Do While ActiveCell.Value <> "" 
    'La sentencia indica que el código del ciclo se ejecutará
    ' mientras el valor de la celda vacía sea algo distinto a un valor vacío ("")

        'Activando las sentencias condicionales
        If_Statements

        'Moviéndome por la columna.
        ActiveCell.Offset(1, 0).Select
        
    Loop
End Sub


'5. CICLOS For Each

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
        'Movi�ndome por la columna.
        ActiveCell.Offset(1, 0).Select
    Next i
    


End Sub
