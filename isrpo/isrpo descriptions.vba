' Example how to write a function description

Sub DescribeOtpusknye()
Dim a As String
Dim b As String
a = "Возвращает количество отпускных"
b = "Общая заработная плата"
c = "Число выходных дней"
Application.MacroOptions _
Macro:="Otpusknye", _
Description:=a, _
ArgumentDescriptions:=Array(b, c)
End Sub
