<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfirmExit
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfirmExit))
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnDoNotExit = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnExit
        '
        resources.ApplyResources(Me.btnExit, "btnExit")
        Me.btnExit.BackColor = System.Drawing.Color.Transparent
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Tag = "TRUE"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnDoNotExit
        '
        resources.ApplyResources(Me.btnDoNotExit, "btnDoNotExit")
        Me.btnDoNotExit.BackColor = System.Drawing.Color.Transparent
        Me.btnDoNotExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnDoNotExit.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.btnDoNotExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDoNotExit.Name = "btnDoNotExit"
        Me.btnDoNotExit.Tag = "FALSE"
        Me.btnDoNotExit.UseVisualStyleBackColor = False
        '
        'frmConfirmExit
        '
        Me.AcceptButton = Me.btnDoNotExit
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.CancelButton = Me.btnDoNotExit
        Me.ControlBox = False
        Me.Controls.Add(Me.btnDoNotExit)
        Me.Controls.Add(Me.btnExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmConfirmExit"
        Me.ShowInTaskbar = False
        Me.TransparencyKey = System.Drawing.Color.Gainsboro
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDoNotExit As System.Windows.Forms.Button
End Class
