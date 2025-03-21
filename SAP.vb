﻿Imports System
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading

Public Module SAP
    Dim path As String
    Public Sub defaultpath()
        path = "C:\Users\paisagu\Mars Inc\Planning South LATAM - Documents\Reportes\Diarios (SAP)"
    End Sub

    Sub AbrirSAP()
        Dim SapGui
        Dim Applic
        Dim connection
        Dim WSHShell

        Try


            Try

                SapGui = GetObject("SAPGUI")
                Applic = SapGui.GetScriptingEngine
                If Applic.Connections.Count > 0 Then
                    Exit Sub
                End If
                connection = Applic.OpenConnection("TRINITY: ERP R/3 Production[SSO]", True)
                Exit Sub
            Catch ex As Exception
                Shell("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus)
                WSHShell = CreateObject("WScript.Shell")
                Do Until WSHShell.AppActivate("SAP Logon ")

                Loop
                WSHShell = Nothing
                SapGui = GetObject("SAPGUI")
                Applic = SapGui.GetScriptingEngine
                connection = Applic.OpenConnection("TRINITY: ERP R/3 Production[SSO]", True)

            End Try
        Catch
            Exit Sub
        End Try


    End Sub

    Sub CerrarSAP()
        Dim SapGui
        Dim Applic
        Dim connection

        SapGui = GetObject("SAPGUI")
        Applic = SapGui.GetScriptingEngine
        connection = Applic.Children(0)

        If Not connection Is Nothing Then
            connection.CloseSession("ses[0]")
            connection = Nothing
        End If

        If Not Applic Is Nothing Then
            Applic.Quit
            Applic = Nothing
        End If

    End Sub
    Sub AtRisk()

        Dim SapGuiAuto As Object
        Dim app As Object
        Dim connection As Object
        Dim session As Object
        Dim connectionNumber As Integer = 0
        Dim sessionNumber As Integer = 0

        Try
            SapGuiAuto = GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            app.HistoryEnabled = False
            connection = app.Children(CInt(connectionNumber))
            If connection.DisabledByServer = True Then Exit Sub
            session = connection.Children(CInt(sessionNumber))
            If session.Info.IsLowSpeedConnection = True Then Exit Sub
        Catch
            Exit Sub
        End Try


        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nzatrisk"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(17)
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]").sendVKey(2)
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "At Risk AR.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
        session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition(81)
        session.findById("wnd[1]/tbar[0]/btn[11]").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]").sendVKey(17)
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
        session.findById("wnd[1]").sendVKey(2)
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "At Risk CH.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition(15)
        session.findById("wnd[1]/tbar[0]/btn[11]").press


        app.HistoryEnabled = True
        Thread.Sleep(6000)
        CerrarExcel("At Risk AR.XLSX")
        Thread.Sleep(4000)
        CerrarExcel("At Risk CH.XLSX")


    End Sub

    Sub Expired()
        Dim SapGuiAuto As Object
        Dim app As Object
        Dim connection As Object
        Dim session As Object

        Dim connectionNumber As Integer = 0
        Dim sessionNumber As Integer = 0

        Try
            SapGuiAuto = GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            app.HistoryEnabled = False
            connection = app.Children(CInt(connectionNumber))
            If connection.DisabledByServer = True Then Exit Sub
            session = connection.Children(CInt(sessionNumber))
            If session.Info.IsLowSpeedConnection = True Then Exit Sub
        Catch
            Exit Sub
        End Try
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nz_stockexpired"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(17)
        session.findById("wnd[1]").sendVKey(8)
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]").sendVKey(2)
        session.findById("wnd[0]/usr/ctxtSP$00004-LOW").text = String.Empty
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/mbar/menu[0]/menu[4]/menu[1]").select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Stock Exprired AR.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
        session.findById("wnd[1]/tbar[0]/btn[11]").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]").sendVKey(17)
        session.findById("wnd[1]").sendVKey(8)
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
        session.findById("wnd[1]").sendVKey(2)
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/mbar/menu[0]/menu[4]/menu[1]").select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Stock Exprired CH.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
        session.findById("wnd[1]/tbar[0]/btn[11]").press


        app.HistoryEnabled = True
        Thread.Sleep(4000)
        CerrarExcel("Stock Exprired AR.XLSX")
        Thread.Sleep(3000)
        CerrarExcel("Stock Exprired CH.XLSX")

    End Sub

    Sub BimReport()
        Dim SapGuiAuto As Object
        Dim app As Object
        Dim connection As Object
        Dim session As Object

        Dim connectionNumber As Integer = 0
        Dim sessionNumber As Integer = 0

        Try
            SapGuiAuto = GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            app.HistoryEnabled = False
            connection = app.Children(CInt(connectionNumber))
            If connection.DisabledByServer = True Then Exit Sub
            session = connection.Children(CInt(sessionNumber))
            If session.Info.IsLowSpeedConnection = True Then Exit Sub
        Catch
            Exit Sub
        End Try
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsd_bim_report"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(17)
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
        session.findById("wnd[1]").sendVKey(2)
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/mbar/menu[0]/menu[4]/menu[1]").select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "BIM Report CH.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
        session.findById("wnd[1]/tbar[0]/btn[11]").press
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsd_bim_report_ar"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(17)
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]").sendVKey(2)
        session.findById("wnd[0]").sendVKey(17)
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
        session.findById("wnd[1]").sendVKey(2)
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/mbar/menu[0]/menu[4]/menu[1]").select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "BIM Report AR.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
        session.findById("wnd[1]/tbar[0]/btn[11]").press

        app.HistoryEnabled = True

        Thread.Sleep(8000)
        CerrarExcel("BIM Report AR.XLSX")
        Thread.Sleep(10000)
        CerrarExcel("BIM Report CH.XLSX")

    End Sub

    Sub Reportes()
        Dim SapGuiAuto As Object
        Dim app As Object
        Dim connection As Object
        Dim session As Object

        Dim connectionNumber As Integer = 0
        Dim sessionNumber As Integer = 0

        Try
            SapGuiAuto = GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            app.HistoryEnabled = False
            connection = app.Children(CInt(connectionNumber))
            If connection.DisabledByServer = True Then Exit Sub
            session = connection.Children(CInt(sessionNumber))
            If session.Info.IsLowSpeedConnection = True Then Exit Sub
        Catch
            Exit Sub
        End Try

        Try
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nzibt_prodreports"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").sendVKey(17)
            session.findById("wnd[1]").sendVKey(8)
            session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
            session.findById("wnd[1]").sendVKey(2)
            session.findById("wnd[0]").sendVKey(8)
            session.findById("wnd[1]/usr/btnBUTTON_1").press
            session.findById("wnd[0]/tbar[1]/btn[43]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Prod Total.XLSX"
            session.findById("wnd[1]/tbar[0]/btn[11]").press

            'session.findById("wnd[0]/tbar[0]/btn[3]").press
            'session.findById("wnd[0]").sendVKey(17)
            'session.findById("wnd[1]").sendVKey(8)
            'session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
            'session.findById("wnd[1]").sendVKey(2)
            'session.findById("wnd[0]").sendVKey(8)
            'session.findById("wnd[1]/usr/btnBUTTON_1").press
            'session.findById("wnd[0]/tbar[1]/btn[43]").press
            'session.findById("wnd[1]/tbar[0]/btn[0]").press
            'session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
            'session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Prod WET.XLSX"
            'session.findById("wnd[1]/tbar[0]/btn[11]").press

            'session.findById("wnd[0]/tbar[0]/btn[3]").press
            'session.findById("wnd[0]").sendVKey(17)
            'session.findById("wnd[1]/usr/txtENAME-LOW").text = "vilagis"
            'session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
            'session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
            'session.findById("wnd[1]").sendVKey(8)
            'session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
            'session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
            'session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
            'session.findById("wnd[0]").sendVKey(8)
            'session.findById("wnd[1]/usr/btnBUTTON_1").press
            'session.findById("wnd[0]/tbar[1]/btn[43]").press
            'session.findById("wnd[1]/tbar[0]/btn[0]").press
            'session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
            'session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Prod C&T.XLSX"
            'session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
            'session.findById("wnd[1]/tbar[0]/btn[11]").press

            app.HistoryEnabled = True
            Thread.Sleep(6000)
            CerrarExcel("Prod Total.XLSX")
            'Thread.Sleep(6000)
            'CerrarExcel("Prod WET.XLSX")
        Catch
            Exit Sub
        End Try

        Thread.Sleep(6000)
        CerrarExcel("Prod Total.XLSX")
        'Thread.Sleep(6000)
        'CerrarExcel("Prod WET.XLSX")
        'Thread.Sleep(8000)
        'CerrarExcel("Prod C&T.XLSX")
    End Sub
    Sub Demanda()
        Dim SapGuiAuto As Object
        Dim app As Object
        Dim connection As Object
        Dim session As Object

        Dim connectionNumber As Integer = 0
        Dim sessionNumber As Integer = 0

        Try
            SapGuiAuto = GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            app.HistoryEnabled = False
            connection = app.Children(CInt(connectionNumber))
            If connection.DisabledByServer = True Then Exit Sub
            session = connection.Children(CInt(sessionNumber))
            If session.Info.IsLowSpeedConnection = True Then Exit Sub
        Catch
            Exit Sub
        End Try

        Try

            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsd_cart_pedidos"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").sendVKey(17)
            session.findById("wnd[1]").sendVKey(8)
            session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").text = DateTime.Today.ToString("ddMMyy")
            session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").setFocus
            session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").caretPosition = 6
            session.findById("wnd[0]").sendVKey(8)
            session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Demanda.XLSX"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
            session.findById("wnd[1]/tbar[0]/btn[11]").press

            app.HistoryEnabled = True

            Thread.Sleep(8000)
            CerrarExcel("Demanda.XLSX")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Catch
        Exit Sub
        End Try

    End Sub
    Sub DemandaSinFecha()
        Dim SapGuiAuto As Object
        Dim app As Object
        Dim connection As Object
        Dim session As Object
        Dim i As Integer
        Dim index As Integer
        Dim rowCount As Integer
        Dim SAPGrid As Object

        Dim connectionNumber As Integer = 0
        Dim sessionNumber As Integer = 0

        Try
            SapGuiAuto = GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            app.HistoryEnabled = False
            connection = app.Children(CInt(connectionNumber))
            If connection.DisabledByServer = True Then Exit Sub
            session = connection.Children(CInt(sessionNumber))
            If session.Info.IsLowSpeedConnection = True Then Exit Sub
        Catch
            Exit Sub
        End Try
        Try
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL06O"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/btnBUTTON6").press
            session.findById("wnd[0]").sendVKey(17)
            session.findById("wnd[1]/usr/txtENAME-LOW").Text = "villafra"
            session.findById("wnd[1]").sendVKey(8)

            session.findById("wnd[0]").sendVKey(8)
            ' session.findById("wnd[0]/usr/lbl[25,5]").SetFocus
            'session.findById("wnd[0]/usr/lbl[25,5]").caretPosition = 2
            'session.findById("wnd[0]/mbar/menu[3]/menu[2]/menu[1]").Select
            session.findById("wnd[0]/tbar[1]/btn[33]").press
            session.findById("wnd[1]/tbar[0]/btn[71]").press
            session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = False
            session.findById("wnd[2]/usr/chkSCAN_STRING-RANGE").selected = True
            'session.findById("wnd[1]/tbar[0]/btn[71]").press
            session.findById("wnd[2]/usr/txtRSYSF-STRING").Text = "FRANCO"
            session.findById("wnd[2]/usr/txtRSYSF-STRING").caretPosition = 10
            session.findById("wnd[2]/tbar[0]/btn[0]").press
            session.findById("wnd[3]/usr/lbl[2,2]").SetFocus
            session.findById("wnd[3]/usr/lbl[2,2]").caretPosition = 7
            session.findById("wnd[3]").sendVKey(2)
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[0]/mbar/menu[0]/menu[5]/menu[1]").Select
            session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Demanda sin Fecha.XLSX"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
            session.findById("wnd[1]/tbar[0]/btn[11]").press

            'session.findById("wnd[0]/tbar[1]/btn[33]").press
            'session.findById("wnd[1]/usr/lbl[1,17]").setFocus
            'session.findById("wnd[1]/usr/lbl[1,17]").caretPosition = 2
            'session.findById("wnd[1]").sendVKey(2)
            'session.findById("wnd[0]/mbar/menu[0]/menu[5]/menu[1]").select
            'session.findById("wnd[1]/tbar[0]/btn[0]").press
            'session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
            'session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Demanda sin Fecha.XLSX"
            'session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
            'session.findById("wnd[1]").sendVKey(11)

            app.HistoryEnabled = True

            Thread.Sleep(8000)
            CerrarExcel("Demanda sin Fecha.XLSX")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub Transitos()
        Dim SapGuiAuto As Object
        Dim app As Object
        Dim connection As Object
        Dim session As Object

        Dim connectionNumber As Integer = 0
        Dim sessionNumber As Integer = 0

        Try
            SapGuiAuto = GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            app.HistoryEnabled = False
            connection = app.Children(CInt(connectionNumber))
            If connection.DisabledByServer = True Then Exit Sub
            session = connection.Children(CInt(sessionNumber))
            If session.Info.IsLowSpeedConnection = True Then Exit Sub
        Catch
            Exit Sub
        End Try

        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb5t"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(17)
        session.findById("wnd[1]").sendVKey(12)
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "ar01"
        session.findById("wnd[0]/usr/ctxtRESWK-LOW").text = "ar06"
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/tbar[1]/btn[43]").press
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Transitos AR06.XLSX"
        session.findById("wnd[1]").sendVKey(11)
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "ar06"
        session.findById("wnd[0]/usr/ctxtRESWK-LOW").text = "ar01"
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/tbar[1]/btn[43]").press
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Transitos AR01.XLSX"
        session.findById("wnd[1]").sendVKey(11)

        app.HistoryEnabled = True

        Thread.Sleep(6000)
        CerrarExcel("Transitos AR06.XLSX")
        Thread.Sleep(8000)
        CerrarExcel("Transitos AR01.XLSX")



    End Sub

    Sub MenuPrincipal()
        Dim SapGuiAuto As Object
        Dim app As Object
        Dim connection As Object
        Dim session As Object

        Dim connectionNumber As Integer = 0
        Dim sessionNumber As Integer = 0

        Try
            SapGuiAuto = GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            app.HistoryEnabled = False
            connection = app.Children(CInt(connectionNumber))
            If connection.DisabledByServer = True Then Exit Sub
            session = connection.Children(CInt(sessionNumber))
            If session.Info.IsLowSpeedConnection = True Then Exit Sub
        Catch
            Exit Sub
        End Try

        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

    End Sub
    Public Sub CerrarExcel(NombreArchivo As String)

        ' Buscar una instancia de Excel en ejecución
        Dim excelApp As Excel.Application = Nothing
        Try
            excelApp = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
        Catch ex As Exception
            ' Manejar la excepción si no se encuentra ninguna instancia de Excel en ejecución
            Console.WriteLine("No se encontró ninguna instancia de Excel en ejecución.")
            Return
        End Try

        Try
            If excelApp.Workbooks.Count = 1 Then
                ' Cerrar completamente la instancia de Excel
                excelApp.Quit()
                Marshal.ReleaseComObject(excelApp)
                excelApp = Nothing
                Return
            End If


            ' Iterar a través de los libros abiertos en Excel
            For Each workbook As Excel.Workbook In excelApp.Workbooks
                ' Verificar si es el libro que deseas cerrar
                If workbook.Name = NombreArchivo Then
                    ' Cerrar el libro sin guardar cambios
                    workbook.Close(SaveChanges:=False)
                    Exit For ' Salir del bucle después de cerrar el libro deseado
                End If
            Next

            ' Liberar recursos
            Marshal.ReleaseComObject(excelApp)
            excelApp = Nothing
        Catch
        End Try
    End Sub

End Module


