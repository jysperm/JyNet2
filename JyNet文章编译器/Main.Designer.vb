<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.ButOK = New System.Windows.Forms.Button()
        Me.BarOut = New System.Windows.Forms.ProgressBar()
        Me.TextOut = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'ButOK
        '
        Me.ButOK.Location = New System.Drawing.Point(396, 129)
        Me.ButOK.Name = "ButOK"
        Me.ButOK.Size = New System.Drawing.Size(119, 41)
        Me.ButOK.TabIndex = 0
        Me.ButOK.Text = "编译"
        Me.ButOK.UseVisualStyleBackColor = True
        '
        'BarOut
        '
        Me.BarOut.Location = New System.Drawing.Point(12, 12)
        Me.BarOut.Name = "BarOut"
        Me.BarOut.Size = New System.Drawing.Size(503, 21)
        Me.BarOut.TabIndex = 1
        '
        'TextOut
        '
        Me.TextOut.Location = New System.Drawing.Point(12, 39)
        Me.TextOut.Multiline = True
        Me.TextOut.Name = "TextOut"
        Me.TextOut.Size = New System.Drawing.Size(254, 131)
        Me.TextOut.TabIndex = 2
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(520, 179)
        Me.Controls.Add(Me.TextOut)
        Me.Controls.Add(Me.BarOut)
        Me.Controls.Add(Me.ButOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "JyNet文章编译器"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButOK As System.Windows.Forms.Button
    Friend WithEvents BarOut As System.Windows.Forms.ProgressBar
    Friend WithEvents TextOut As System.Windows.Forms.TextBox

End Class
