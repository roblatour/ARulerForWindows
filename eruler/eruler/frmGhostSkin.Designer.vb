<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGhostSkin
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
        Me.VerticalRuler = New System.Windows.Forms.Panel()
        Me.HorizontalRuler = New System.Windows.Forms.Panel()
        Me.SuspendLayout()
        '
        'VerticalRuler
        '
        Me.VerticalRuler.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.VerticalRuler.AutoSize = True
        Me.VerticalRuler.BackColor = System.Drawing.Color.LavenderBlush
        Me.VerticalRuler.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.VerticalRuler.Location = New System.Drawing.Point(370, 12)
        Me.VerticalRuler.Name = "VerticalRuler"
        Me.VerticalRuler.Size = New System.Drawing.Size(100, 350)
        Me.VerticalRuler.TabIndex = 8
        '
        'HorizontalRuler
        '
        Me.HorizontalRuler.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.HorizontalRuler.AutoSize = True
        Me.HorizontalRuler.BackColor = System.Drawing.Color.LavenderBlush
        Me.HorizontalRuler.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.HorizontalRuler.Location = New System.Drawing.Point(14, 343)
        Me.HorizontalRuler.Name = "HorizontalRuler"
        Me.HorizontalRuler.Size = New System.Drawing.Size(350, 100)
        Me.HorizontalRuler.TabIndex = 7
        '
        'frmGhostSkin
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(474, 446)
        Me.Controls.Add(Me.VerticalRuler)
        Me.Controls.Add(Me.HorizontalRuler)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmGhostSkin"
        Me.ShowInTaskbar = False
        Me.Text = "frmGhostSkin"
        Me.TransparencyKey = System.Drawing.Color.LavenderBlush
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents HorizontalRuler As System.Windows.Forms.Panel
    Friend WithEvents VerticalRuler As Panel
End Class
