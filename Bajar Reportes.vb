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
        CerrarSAP()
        Me.Close()
    End Sub

    Private Sub Bajar_Reportes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        IniciarToolTips()
        AbrirSAP()
        Me.Activate()
        defaultpath()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub btnBajar_Click(sender As Object, e As EventArgs) Handles btnBajar.Click
        'AtRisk()
        BimReport()
        Reportes()
        Demanda()
        DemandaSinFecha()
        Expired()
        Transitos()
        MenuPrincipal()
        CerrarSAP()
        Me.Close()
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

    Public Sub IniciarToolTips()
        TTAtRisk.SetToolTip(btnAtRisk, "Variantes: En Riesgo para Argentina, En Riesgo CH para Chile. ")
        TTStockExpired.SetToolTip(btnExpired, "Variantes: STOCK EXP ARG para Argentina, STOCK EXP CH para Chile.")
        TTBimReport.SetToolTip(btnBIMReport, "Variantes: BIMREPORT para Argentina en la transacción BIM Report Argentina, BIM REPORT CHI para Chile en la transacción BIM Report")
        TTDemanda.SetToolTip(btnDemanda, "Variantes: Para Demanda ON Hand y Demanda sin Fecha, la unica variante en el usuario PAISAGU.")
        TTProduction.SetToolTip(btnReportes, "Variantes: PRODU MER para Produccion Total, PRODU WET para Producción de WET")
        TTTransitos.SetToolTip(btnTransitos, "Para Transitos AR06.xlsx sólo se coloca centro de AR01 a AR06, y para tránsitos AR01.xlsx se coloca centro de AR06 a AR01")
    End Sub

End Class