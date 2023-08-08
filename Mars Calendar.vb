Imports Microsoft.Office.Interop.Excel
Imports System.Security.Cryptography

Module Mars_Calendar
    Function MarsPeriod(FechaInicio As Date, FechaCalculo As Date, Opcion As Integer) As String
        Dim Periodo, semanas, RestarSemanas, semana As Double
        Dim result As String
        Dim dia As Integer

        Periodo = StringPadLeft(Math.Ceiling((FechaCalculo.ToOADate() - FechaInicio.ToOADate()) * 13 / 364), 2, "0")
        semanas = Math.Ceiling((FechaCalculo.ToOADate() - FechaInicio.ToOADate()) / 7)
        RestarSemanas = (Periodo - 1) * 4
        semana = semanas - RestarSemanas
        dia = Weekday(FechaCalculo, 1)
        Select Case Opcion
            Case 1
                result = "P" & Periodo & "W" & semana & "D" & dia
            Case 2
                result = "P" & Periodo & "W" & semana
            Case 3
                result = "P" & Periodo
            Case 4
                result = "W" & semana
            Case 5
                If semanas < 10 Then
                    result = "0" & semanas
                Else
                    result = semanas
                End If
            Case Else
                result = ""
        End Select
        MarsPeriod = result

    End Function

    Public Function StringPadLeft(Expression As String, width As Integer, Optional character As String = " ")
        StringPadLeft = Right(StrDup(width, character) & Expression, width)
    End Function

End Module
