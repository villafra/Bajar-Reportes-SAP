Public Class Bajar_Reportes
    Private Sub btnAtRisk_Click(sender As Object, e As EventArgs) Handles btnAtRisk.Click
        AtRisk()
        MenuPrincipal()
    End Sub

    Private Sub btnExpired_Click(sender As Object, e As EventArgs) Handles btnExpired.Click
        Expired()
        MenuPrincipal()
    End Sub

    Private Sub btnBIMReport_Click(sender As Object, e As EventArgs) Handles btnBIMReport.Click
        BimReport()
        MenuPrincipal()
    End Sub

    Private Sub btnReportes_Click(sender As Object, e As EventArgs) Handles btnReportes.Click
        Reportes()
        MenuPrincipal()
    End Sub

    Private Sub Bajar_Reportes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AbrirSAP()
        Me.Activate()
        defaultpath()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub btnBajar_Click(sender As Object, e As EventArgs) Handles btnBajar.Click
        AtRisk()
        BimReport()
        Reportes()
        Demanda()
        DemandaSinFecha()
        Expired()
        Transitos()
        MenuPrincipal()
    End Sub

    Private Sub btnDemanda_Click(sender As Object, e As EventArgs) Handles btnDemanda.Click
        DemandaSinFecha()
        Demanda()
        MenuPrincipal()
    End Sub

    Private Sub btnTransitos_Click(sender As Object, e As EventArgs) Handles btnTransitos.Click
        Transitos()
        MenuPrincipal()
    End Sub
End Class