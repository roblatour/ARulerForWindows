<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRuler
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRuler))
        Me.MagnifyViewPanel = New System.Windows.Forms.Panel()
        Me.RulerPanel = New System.Windows.Forms.Panel()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.PilotFishPanel = New System.Windows.Forms.Panel()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'MagnifyViewPanel
        '
        resources.ApplyResources(Me.MagnifyViewPanel, "MagnifyViewPanel")
        Me.MagnifyViewPanel.Name = "MagnifyViewPanel"
        '
        'RulerPanel
        '
        Me.RulerPanel.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.RulerPanel, "RulerPanel")
        Me.RulerPanel.Name = "RulerPanel"
        '
        'ToolTip1
        '
        Me.ToolTip1.AutomaticDelay = 300
        Me.ToolTip1.AutoPopDelay = 7000
        Me.ToolTip1.InitialDelay = 10
        Me.ToolTip1.ReshowDelay = 10
        '
        'PilotFishPanel
        '
        Me.PilotFishPanel.BackColor = System.Drawing.Color.Red
        resources.ApplyResources(Me.PilotFishPanel, "PilotFishPanel")
        Me.PilotFishPanel.Name = "PilotFishPanel"
        '
        'Timer1
        '
        Me.Timer1.Interval = 1500
        '
        'Timer2
        '
        '
        'frmRuler
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        resources.ApplyResources(Me, "$this")
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(250, Byte), Integer), CType(CType(250, Byte), Integer))
        Me.Controls.Add(Me.PilotFishPanel)
        Me.Controls.Add(Me.RulerPanel)
        Me.Controls.Add(Me.MagnifyViewPanel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmRuler"
        Me.TransparencyKey = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(250, Byte), Integer), CType(CType(250, Byte), Integer))
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents MagnifyViewPanel As System.Windows.Forms.Panel
    Friend WithEvents RulerPanel As System.Windows.Forms.Panel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents PilotFishPanel As System.Windows.Forms.Panel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer2 As Timer
End Class
