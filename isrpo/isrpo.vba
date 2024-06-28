' Exam docx


' ================================================
' RemoveSpaces (All version)
Function REMOVESPACES(cell As Range) As String
    Dim CellLength As Long
    Dim Temp As String
    Dim i As Long
    CellLength = Len(cell)
    Temp = ""
    For i = 1 To CellLength
        Character = Mid(cell, i, 1)
        If Character <> Chr(32) Then Temp = Temp & Character
    Next i
    REMOVESPACES = Temp
End Function

Function REMOVESPACES2(cell As Range) As String
REMOVESPACES2 = Replace(cell, " ", "")
End Function
' ================================================


' ================================================
' With Formulas
Option Explicit
' CellFormula
Function CELLFORMULA(cell As Range) As String
    Application.Volatile True
    If cell.Range("A1").HasFormula Then
        CELLFORMULA = cell.Range("A1").Formula
    Else
        CELLFORMULA = ""
    End If
End Function
' =================================================


' =================================================
'StatFunction
Function StatFunction(range As Range, op As String) As Double
    Select Case range
    Case "СРЗНАЧ"
        StatFunction = Application.WorksheetFunction.Average(range)
    Case "СЧЁТ"
        StatFunction = Application.WorksheetFunction.Count(range)
    Case "МАКС"
        StatFunction = Application.WorksheetFunction.Max(range)
    Case "МЕДИАНА"
        StatFunction = Application.WorksheetFunction.Median(range)
    Case "МИН"
        StatFunction = Application.WorksheetFunction.Min(range)
    Case "МОДА"
        StatFunction = Application.WorksheetFunction.Mode(range)
    Case Else
        StatFunction = "Сорри, неверно указан аргумент, попробуйте ещё раз"
    End Select
' =================================================

' =================================================
' Commission
Public Function Commission(summ As Double, percent As Double)
Commission = summ / 100 * percent
End Function
' =================================================

' =================================================
' CellHasFormula
Function CELLHASFORMULA(cell As Range) As Boolean
    Application.Volatile True
    CELLHASFORMULA = cell.Range("A1").HasFormula
End Function

' CellHidden
Function CELLISHIDDEN(cell As Range) As Boolean
    Application.Volatile True
    Dim UpperLeft As Range
    Set UpperLeft = cell.Range("A1")
    CELLISHIDDEN = UpperLeft.EntireRow.Hidden Or _
        UpperLeft.EntireColumn.Hidden
End Function
' =================================================

' =================================================
' Acronim
Function Acronim(a As String) As String
Dim final As String
Dim check As Boolean
check = True
final = ""
For i = 0 To Len(a)
    If check Then
        final = final & a.Chars(i)
        check = False
    End If

    If a.Chars(i) == " " Then
        check = True
Next i
Acronim = final
End Function
' =================================================

' =================================================
' DayOfWeek
Public Function DayOfWeek(number As Integer) As String
Dim result As String
Select Case number
Case 1
result = "Понедельник"
Case 2
result = "Вторник"
Case 3
result = "Среда"
Case 4
result = "Четверг"
Case 5
result = "Пятница"
Case 6
result = "Суббота"
Case 7
result = "Воскресенье"
Case Else
result = "Неверное значение"
End Select
DayOfWeek = result
End Function
' =================================================

' =================================================
'Otpusknie
Public Function Otpusknye(summZp As Long, holidays As Long) As Long
If IsNumeric(holidays) = False Or IsNumeric(summZp) = False Then
Otpusknye = "Введены нечисловые данные"
Exit Function
ElseIf holidays <= 0 Or summZp <= 0 Then
Otpusknye = "Отрицательное число или 0"
Exit Function
Else
Otpusknye = summZp * 24 / (365 - holidays)
End If
End Function
' =================================================

' =================================================
' CaloriesPerDay
Public Function CaloriesPerDay(sex As String, age As Integer, weight As Integer, height As Integer) As Integer
If sex = "женский" Then
CaloriesPerDay = 10 * weight + 6.25 * height - 5 * age - 161
ElseIf sex = "мужской" Then
CaloriesPerDay = 10 * weight + 6.25 * height - 5 * age + 5
Else: CaloriesPerDay = 0
End If
End Function
' =================================================

' =================================================
' SquareEquation
Public Function SquareEquation(a As Integer, b As Integer, c As Integer) As String
Dim answer1 As String
Dim answer2 As String
If a = 0 Then
answer1 = "Единственный корень - "
SquareEquation = answer1 & "(" & -c / b & ")"
ElseIf c = 0 Then
answer1 = "Единственный корень - "
SquareEquation = answer1 & "(" & -b / a & ")"
ElseIf b = 0 And c < 0 Then
answer1 = "Единственный корень - "
SquareEquation = answer1 & "(" & Sqr(a / c) & ")"
ElseIf b ^ 2 - 4 * a * c >= 0 Then
answer1 = "Первый корень - "
answer2 = "Второй корень - "
SquareEquation = answer1 & "(" & (-b + Sqr(b ^ 2 - 4 * a * c)) / (2 * a) & ")" & "; " & _
answer2 & "(" & (-b - Sqr(b ^ 2 - 4 * a * c)) / (2 * a) & ")"
Else:
SquareEquation = "Решений нет"
End If
End Function
' =================================================
