<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Bajar_Reportes
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Bajar_Reportes))
        Me.btnAtRisk = New System.Windows.Forms.Button()
        Me.btnExpired = New System.Windows.Forms.Button()
        Me.btnBIMReport = New System.Windows.Forms.Button()
        Me.btnReportes = New System.Windows.Forms.Button()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBajar = New System.Windows.Forms.Button()
        Me.btnDemanda = New System.Windows.Forms.Button()
        Me.btnTransitos = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnAtRisk
        '
        Me.btnAtRisk.Location = New System.Drawing.Point(18, 159)
        Me.btnAtRisk.Name = "btnAtRisk"
        Me.btnAtRisk.Size = New System.Drawing.Size(185, 23)
        Me.btnAtRisk.TabIndex = 0
        Me.btnAtRisk.Text = "At Risk"
        Me.btnAtRisk.UseVisualStyleBackColor = True
        '
        'btnExpired
        '
        Me.btnExpired.Location = New System.Drawing.Point(214, 159)
        Me.btnExpired.Name = "btnExpired"
        Me.btnExpired.Size = New System.Drawing.Size(185, 23)
        Me.btnExpired.TabIndex = 1
        Me.btnExpired.Text = "Stock Expired"
        Me.btnExpired.UseVisualStyleBackColor = True
        '
        'btnBIMReport
        '
        Me.btnBIMReport.Location = New System.Drawing.Point(410, 159)
        Me.btnBIMReport.Name = "btnBIMReport"
        Me.btnBIMReport.Size = New System.Drawing.Size(185, 23)
        Me.btnBIMReport.TabIndex = 2
        Me.btnBIMReport.Text = "BIM Report"
        Me.btnBIMReport.UseVisualStyleBackColor = True
        '
        'btnReportes
        '
        Me.btnReportes.Location = New System.Drawing.Point(606, 159)
        Me.btnReportes.Name = "btnReportes"
        Me.btnReportes.Size = New System.Drawing.Size(185, 23)
        Me.btnReportes.TabIndex = 3
        Me.btnReportes.Text = "Reportes Producción"
        Me.btnReportes.UseVisualStyleBackColor = True
        '
        'btnSalir
        '
        Me.btnSalir.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnSalir.Location = New System.Drawing.Point(898, 2)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(103, 23)
        Me.btnSalir.TabIndex = 4
        Me.btnSalir.Text = "Salir"
        Me.btnSalir.UseVisualStyleBackColor = True
        Me.btnSalir.Visible = False
        '
        'btnBajar
        '
        Me.btnBajar.Location = New System.Drawing.Point(410, 92)
        Me.btnBajar.Name = "btnBajar"
        Me.btnBajar.Size = New System.Drawing.Size(185, 23)
        Me.btnBajar.TabIndex = 5
        Me.btnBajar.Text = "Bajar Todo"
        Me.btnBajar.UseVisualStyleBackColor = True
        '
        'btnDemanda
        '
        Me.btnDemanda.Location = New System.Drawing.Point(802, 159)
        Me.btnDemanda.Name = "btnDemanda"
        Me.btnDemanda.Size = New System.Drawing.Size(185, 23)
        Me.btnDemanda.TabIndex = 6
        Me.btnDemanda.Text = "Demanda"
        Me.btnDemanda.UseVisualStyleBackColor = True
        '
        'btnTransitos
        '
        Me.btnTransitos.Location = New System.Drawing.Point(18, 12)
        Me.btnTransitos.Name = "btnTransitos"
        Me.btnTransitos.Size = New System.Drawing.Size(185, 23)
        Me.btnTransitos.TabIndex = 7
        Me.btnTransitos.Text = "Tránsitos"
        Me.btnTransitos.UseVisualStyleBackColor = True
        '
        'Bajar_Reportes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.Bajar_Reportes.My.Resources.Resources.hacker_g7f2ae809e_1280
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.CancelButton = Me.btnSalir
        Me.ClientSize = New System.Drawing.Size(1004, 206)
        Me.Controls.Add(Me.btnTransitos)
        Me.Controls.Add(Me.btnDemanda)
        Me.Controls.Add(Me.btnBajar)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.btnReportes)
        Me.Controls.Add(Me.btnBIMReport)
        Me.Controls.Add(Me.btnExpired)
        Me.Controls.Add(Me.btnAtRisk)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Bajar_Reportes"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Bajar_Reportes"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnAtRisk As Button
    Friend WithEvents btnExpired As Button
    Friend WithEvents btnBIMReport As Button
    Friend WithEvents btnReportes As Button
    Friend WithEvents btnSalir As Button
    Friend WithEvents btnBajar As Button
    Friend WithEvents btnDemanda As Button
    Friend WithEvents btnTransitos As Button
End Class
