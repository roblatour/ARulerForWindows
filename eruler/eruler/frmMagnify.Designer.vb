<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMagnify
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMagnify))
        Me.rbNormal = New System.Windows.Forms.RadioButton()
        Me.rb2x = New System.Windows.Forms.RadioButton()
        Me.rb3x = New System.Windows.Forms.RadioButton()
        Me.rb4x = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'rbNormal
        '
        resources.ApplyResources(Me.rbNormal, "rbNormal")
        Me.rbNormal.BackColor = System.Drawing.Color.Transparent
        Me.rbNormal.Name = "rbNormal"
        Me.rbNormal.Tag = "1"
        Me.rbNormal.UseVisualStyleBackColor = False
        '
        'rb2x
        '
        resources.ApplyResources(Me.rb2x, "rb2x")
        Me.rb2x.BackColor = System.Drawing.Color.Transparent
        Me.rb2x.Name = "rb2x"
        Me.rb2x.Tag = "2"
        Me.rb2x.UseVisualStyleBackColor = False
        '
        'rb3x
        '
        resources.ApplyResources(Me.rb3x, "rb3x")
        Me.rb3x.BackColor = System.Drawing.Color.Transparent
        Me.rb3x.Name = "rb3x"
        Me.rb3x.Tag = "3"
        Me.rb3x.UseVisualStyleBackColor = False
        '
        'rb4x
        '
        resources.ApplyResources(Me.rb4x, "rb4x")
        Me.rb4x.BackColor = System.Drawing.Color.Transparent
        Me.rb4x.Name = "rb4x"
        Me.rb4x.Tag = "4"
        Me.rb4x.UseVisualStyleBackColor = False
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Name = "Label1"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Name = "Label2"
        '
        'btnOK
        '
        resources.ApplyResources(Me.btnOK, "btnOK")
        Me.btnOK.Name = "btnOK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'frmMagnify
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        resources.ApplyResources(Me, "$this")
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ControlBox = False
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.rb4x)
        Me.Controls.Add(Me.rb3x)
        Me.Controls.Add(Me.rb2x)
        Me.Controls.Add(Me.rbNormal)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmMagnify"
        Me.ShowInTaskbar = False
        Me.TransparencyKey = System.Drawing.Color.Gainsboro
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rbNormal As System.Windows.Forms.RadioButton
    Friend WithEvents rb2x As System.Windows.Forms.RadioButton
    Friend WithEvents rb3x As System.Windows.Forms.RadioButton
    Friend WithEvents rb4x As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
