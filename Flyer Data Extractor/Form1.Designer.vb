<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmFlyerDataExtrarctor
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.BtnExtractData = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'BtnExtractData
        '
        Me.BtnExtractData.Location = New System.Drawing.Point(12, 12)
        Me.BtnExtractData.Name = "BtnExtractData"
        Me.BtnExtractData.Size = New System.Drawing.Size(75, 23)
        Me.BtnExtractData.TabIndex = 0
        Me.BtnExtractData.Text = "Extract Data"
        Me.BtnExtractData.UseVisualStyleBackColor = True
        '
        'FrmFlyerDataExtrarctor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(422, 261)
        Me.Controls.Add(Me.BtnExtractData)
        Me.Name = "FrmFlyerDataExtrarctor"
        Me.Text = "Flyer Data Extrarctor"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents BtnExtractData As Button
End Class
