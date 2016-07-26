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
        Me.TxtWebsiteUrl = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblInfo = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'BtnExtractData
        '
        Me.BtnExtractData.Location = New System.Drawing.Point(168, 39)
        Me.BtnExtractData.Name = "BtnExtractData"
        Me.BtnExtractData.Size = New System.Drawing.Size(75, 23)
        Me.BtnExtractData.TabIndex = 0
        Me.BtnExtractData.Text = "Extract Data"
        Me.BtnExtractData.UseVisualStyleBackColor = True
        '
        'TxtWebsiteUrl
        '
        Me.TxtWebsiteUrl.Location = New System.Drawing.Point(91, 13)
        Me.TxtWebsiteUrl.Name = "TxtWebsiteUrl"
        Me.TxtWebsiteUrl.Size = New System.Drawing.Size(319, 20)
        Me.TxtWebsiteUrl.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Website URL:"
        '
        'LblInfo
        '
        Me.LblInfo.AutoSize = True
        Me.LblInfo.Location = New System.Drawing.Point(148, 65)
        Me.LblInfo.Name = "LblInfo"
        Me.LblInfo.Size = New System.Drawing.Size(39, 13)
        Me.LblInfo.TabIndex = 3
        Me.LblInfo.Text = "LblInfo"
        '
        'FrmFlyerDataExtrarctor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(422, 261)
        Me.Controls.Add(Me.LblInfo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtWebsiteUrl)
        Me.Controls.Add(Me.BtnExtractData)
        Me.Name = "FrmFlyerDataExtrarctor"
        Me.Text = "Flyer Data Extractor"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BtnExtractData As Button
    Friend WithEvents TxtWebsiteUrl As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents LblInfo As Label
End Class
