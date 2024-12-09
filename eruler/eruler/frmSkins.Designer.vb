<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSkins
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSkins))
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btnLineThirds = New System.Windows.Forms.Button()
        Me.btnLineGoldenRatio = New System.Windows.Forms.Button()
        Me.cbEnabled = New System.Windows.Forms.CheckBox()
        Me.btnLineMidpoint = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.tbRotateHorizontal = New System.Windows.Forms.TrackBar()
        Me.cbStretch = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.tbRotateVertical = New System.Windows.Forms.TrackBar()
        Me.btnDefaults2 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tbTransparency = New System.Windows.Forms.TrackBar()
        Me.btnBackground = New System.Windows.Forms.Button()
        Me.btnLineBoxes = New System.Windows.Forms.Button()
        Me.btnLineTicksAndSides = New System.Windows.Forms.Button()
        Me.btnLineLength = New System.Windows.Forms.Button()
        Me.btnNumbers = New System.Windows.Forms.Button()
        Me.btnLineMeasuring = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.btnGetMoreFreeSkinsOnline = New System.Windows.Forms.Button()
        Me.btnLocate = New System.Windows.Forms.Button()
        Me.btnDefaultSkins = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.cbSkinNames = New System.Windows.Forms.ComboBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.HorizontalRuler = New System.Windows.Forms.Panel()
        Me.VerticalRuler = New System.Windows.Forms.Panel()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.tbRotateHorizontal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.tbRotateVertical, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.tbTransparency, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.White
        Me.GroupBox3.Controls.Add(Me.btnLineThirds)
        Me.GroupBox3.Controls.Add(Me.btnLineGoldenRatio)
        Me.GroupBox3.Controls.Add(Me.cbEnabled)
        Me.GroupBox3.Controls.Add(Me.btnLineMidpoint)
        Me.GroupBox3.Controls.Add(Me.GroupBox4)
        Me.GroupBox3.Controls.Add(Me.cbStretch)
        Me.GroupBox3.Controls.Add(Me.GroupBox2)
        Me.GroupBox3.Controls.Add(Me.btnDefaults2)
        Me.GroupBox3.Controls.Add(Me.GroupBox1)
        Me.GroupBox3.Controls.Add(Me.btnBackground)
        Me.GroupBox3.Controls.Add(Me.btnLineBoxes)
        Me.GroupBox3.Controls.Add(Me.btnLineTicksAndSides)
        Me.GroupBox3.Controls.Add(Me.btnLineLength)
        Me.GroupBox3.Controls.Add(Me.btnNumbers)
        Me.GroupBox3.Controls.Add(Me.btnLineMeasuring)
        resources.ApplyResources(Me.GroupBox3, "GroupBox3")
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.TabStop = False
        '
        'btnLineThirds
        '
        Me.btnLineThirds.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnLineThirds, "btnLineThirds")
        Me.btnLineThirds.Name = "btnLineThirds"
        Me.ToolTip1.SetToolTip(Me.btnLineThirds, resources.GetString("btnLineThirds.ToolTip"))
        Me.btnLineThirds.UseVisualStyleBackColor = False
        '
        'btnLineGoldenRatio
        '
        Me.btnLineGoldenRatio.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnLineGoldenRatio, "btnLineGoldenRatio")
        Me.btnLineGoldenRatio.Name = "btnLineGoldenRatio"
        Me.ToolTip1.SetToolTip(Me.btnLineGoldenRatio, resources.GetString("btnLineGoldenRatio.ToolTip"))
        Me.btnLineGoldenRatio.UseVisualStyleBackColor = False
        '
        'cbEnabled
        '
        resources.ApplyResources(Me.cbEnabled, "cbEnabled")
        Me.cbEnabled.Name = "cbEnabled"
        Me.ToolTip1.SetToolTip(Me.cbEnabled, resources.GetString("cbEnabled.ToolTip"))
        Me.cbEnabled.UseVisualStyleBackColor = True
        '
        'btnLineMidpoint
        '
        Me.btnLineMidpoint.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnLineMidpoint, "btnLineMidpoint")
        Me.btnLineMidpoint.Name = "btnLineMidpoint"
        Me.ToolTip1.SetToolTip(Me.btnLineMidpoint, resources.GetString("btnLineMidpoint.ToolTip"))
        Me.btnLineMidpoint.UseVisualStyleBackColor = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.tbRotateHorizontal)
        resources.ApplyResources(Me.GroupBox4, "GroupBox4")
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.TabStop = False
        '
        'tbRotateHorizontal
        '
        resources.ApplyResources(Me.tbRotateHorizontal, "tbRotateHorizontal")
        Me.tbRotateHorizontal.Maximum = 4
        Me.tbRotateHorizontal.Minimum = 1
        Me.tbRotateHorizontal.Name = "tbRotateHorizontal"
        Me.tbRotateHorizontal.TickStyle = System.Windows.Forms.TickStyle.Both
        Me.ToolTip1.SetToolTip(Me.tbRotateHorizontal, resources.GetString("tbRotateHorizontal.ToolTip"))
        Me.tbRotateHorizontal.Value = 1
        '
        'cbStretch
        '
        resources.ApplyResources(Me.cbStretch, "cbStretch")
        Me.cbStretch.Name = "cbStretch"
        Me.ToolTip1.SetToolTip(Me.cbStretch, resources.GetString("cbStretch.ToolTip"))
        Me.cbStretch.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.tbRotateVertical)
        resources.ApplyResources(Me.GroupBox2, "GroupBox2")
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.TabStop = False
        '
        'tbRotateVertical
        '
        resources.ApplyResources(Me.tbRotateVertical, "tbRotateVertical")
        Me.tbRotateVertical.Maximum = 4
        Me.tbRotateVertical.Minimum = 1
        Me.tbRotateVertical.Name = "tbRotateVertical"
        Me.tbRotateVertical.TickStyle = System.Windows.Forms.TickStyle.Both
        Me.ToolTip1.SetToolTip(Me.tbRotateVertical, resources.GetString("tbRotateVertical.ToolTip"))
        Me.tbRotateVertical.Value = 1
        '
        'btnDefaults2
        '
        Me.btnDefaults2.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnDefaults2, "btnDefaults2")
        Me.btnDefaults2.Name = "btnDefaults2"
        Me.ToolTip1.SetToolTip(Me.btnDefaults2, resources.GetString("btnDefaults2.ToolTip"))
        Me.btnDefaults2.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.tbTransparency)
        resources.ApplyResources(Me.GroupBox1, "GroupBox1")
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.TabStop = False
        '
        'tbTransparency
        '
        resources.ApplyResources(Me.tbTransparency, "tbTransparency")
        Me.tbTransparency.Maximum = 100
        Me.tbTransparency.Name = "tbTransparency"
        Me.tbTransparency.TickFrequency = 10
        Me.tbTransparency.TickStyle = System.Windows.Forms.TickStyle.Both
        Me.ToolTip1.SetToolTip(Me.tbTransparency, resources.GetString("tbTransparency.ToolTip"))
        Me.tbTransparency.Value = 100
        '
        'btnBackground
        '
        Me.btnBackground.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnBackground, "btnBackground")
        Me.btnBackground.Name = "btnBackground"
        Me.ToolTip1.SetToolTip(Me.btnBackground, resources.GetString("btnBackground.ToolTip"))
        Me.btnBackground.UseVisualStyleBackColor = False
        '
        'btnLineBoxes
        '
        Me.btnLineBoxes.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnLineBoxes, "btnLineBoxes")
        Me.btnLineBoxes.Name = "btnLineBoxes"
        Me.ToolTip1.SetToolTip(Me.btnLineBoxes, resources.GetString("btnLineBoxes.ToolTip"))
        Me.btnLineBoxes.UseVisualStyleBackColor = False
        '
        'btnLineTicksAndSides
        '
        Me.btnLineTicksAndSides.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnLineTicksAndSides, "btnLineTicksAndSides")
        Me.btnLineTicksAndSides.Name = "btnLineTicksAndSides"
        Me.ToolTip1.SetToolTip(Me.btnLineTicksAndSides, resources.GetString("btnLineTicksAndSides.ToolTip"))
        Me.btnLineTicksAndSides.UseVisualStyleBackColor = False
        '
        'btnLineLength
        '
        Me.btnLineLength.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnLineLength, "btnLineLength")
        Me.btnLineLength.Name = "btnLineLength"
        Me.ToolTip1.SetToolTip(Me.btnLineLength, resources.GetString("btnLineLength.ToolTip"))
        Me.btnLineLength.UseVisualStyleBackColor = False
        '
        'btnNumbers
        '
        Me.btnNumbers.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnNumbers, "btnNumbers")
        Me.btnNumbers.Name = "btnNumbers"
        Me.ToolTip1.SetToolTip(Me.btnNumbers, resources.GetString("btnNumbers.ToolTip"))
        Me.btnNumbers.UseVisualStyleBackColor = False
        '
        'btnLineMeasuring
        '
        Me.btnLineMeasuring.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnLineMeasuring, "btnLineMeasuring")
        Me.btnLineMeasuring.Name = "btnLineMeasuring"
        Me.ToolTip1.SetToolTip(Me.btnLineMeasuring, resources.GetString("btnLineMeasuring.ToolTip"))
        Me.btnLineMeasuring.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        resources.ApplyResources(Me.btnOK, "btnOK")
        Me.btnOK.BackColor = System.Drawing.Color.White
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnOK.ForeColor = System.Drawing.Color.Black
        Me.btnOK.Name = "btnOK"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.Color.White
        Me.GroupBox5.Controls.Add(Me.btnGetMoreFreeSkinsOnline)
        Me.GroupBox5.Controls.Add(Me.btnLocate)
        Me.GroupBox5.Controls.Add(Me.btnDefaultSkins)
        Me.GroupBox5.Controls.Add(Me.btnRemove)
        Me.GroupBox5.Controls.Add(Me.btnAdd)
        Me.GroupBox5.Controls.Add(Me.cbSkinNames)
        resources.ApplyResources(Me.GroupBox5, "GroupBox5")
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.TabStop = False
        '
        'btnGetMoreFreeSkinsOnline
        '
        Me.btnGetMoreFreeSkinsOnline.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnGetMoreFreeSkinsOnline, "btnGetMoreFreeSkinsOnline")
        Me.btnGetMoreFreeSkinsOnline.Name = "btnGetMoreFreeSkinsOnline"
        Me.ToolTip1.SetToolTip(Me.btnGetMoreFreeSkinsOnline, resources.GetString("btnGetMoreFreeSkinsOnline.ToolTip"))
        Me.btnGetMoreFreeSkinsOnline.UseVisualStyleBackColor = False
        '
        'btnLocate
        '
        Me.btnLocate.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnLocate, "btnLocate")
        Me.btnLocate.Name = "btnLocate"
        Me.ToolTip1.SetToolTip(Me.btnLocate, resources.GetString("btnLocate.ToolTip"))
        Me.btnLocate.UseVisualStyleBackColor = False
        '
        'btnDefaultSkins
        '
        Me.btnDefaultSkins.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnDefaultSkins, "btnDefaultSkins")
        Me.btnDefaultSkins.Name = "btnDefaultSkins"
        Me.ToolTip1.SetToolTip(Me.btnDefaultSkins, resources.GetString("btnDefaultSkins.ToolTip"))
        Me.btnDefaultSkins.UseVisualStyleBackColor = False
        '
        'btnRemove
        '
        Me.btnRemove.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnRemove, "btnRemove")
        Me.btnRemove.Name = "btnRemove"
        Me.ToolTip1.SetToolTip(Me.btnRemove, resources.GetString("btnRemove.ToolTip"))
        Me.btnRemove.UseVisualStyleBackColor = False
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnAdd, "btnAdd")
        Me.btnAdd.Name = "btnAdd"
        Me.ToolTip1.SetToolTip(Me.btnAdd, resources.GetString("btnAdd.ToolTip"))
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'cbSkinNames
        '
        Me.cbSkinNames.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbSkinNames.FormattingEnabled = True
        resources.ApplyResources(Me.cbSkinNames, "cbSkinNames")
        Me.cbSkinNames.Name = "cbSkinNames"
        Me.cbSkinNames.Sorted = True
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.White
        resources.ApplyResources(Me.btnSave, "btnSave")
        Me.btnSave.Name = "btnSave"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'HorizontalRuler
        '
        resources.ApplyResources(Me.HorizontalRuler, "HorizontalRuler")
        Me.HorizontalRuler.BackColor = System.Drawing.Color.LavenderBlush
        Me.HorizontalRuler.Name = "HorizontalRuler"
        '
        'VerticalRuler
        '
        resources.ApplyResources(Me.VerticalRuler, "VerticalRuler")
        Me.VerticalRuler.BackColor = System.Drawing.Color.LavenderBlush
        Me.VerticalRuler.Name = "VerticalRuler"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        resources.ApplyResources(Me.ContextMenuStrip1, "ContextMenuStrip1")
        '
        'frmSkins
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        resources.ApplyResources(Me, "$this")
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.btnOK
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.VerticalRuler)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.HorizontalRuler)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSkins"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.TransparencyKey = System.Drawing.Color.LavenderBlush
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.tbRotateHorizontal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.tbRotateVertical, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.tbTransparency, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnNumbers As System.Windows.Forms.Button
    Friend WithEvents btnLineMidpoint As System.Windows.Forms.Button
    Friend WithEvents btnLineMeasuring As System.Windows.Forms.Button
    Friend WithEvents btnLineBoxes As System.Windows.Forms.Button
    Friend WithEvents btnBackground As System.Windows.Forms.Button
    Friend WithEvents btnLineLength As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents cbSkinNames As System.Windows.Forms.ComboBox
    Friend WithEvents HorizontalRuler As System.Windows.Forms.Panel
    Friend WithEvents VerticalRuler As System.Windows.Forms.Panel
    Friend WithEvents btnLineTicksAndSides As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents tbTransparency As System.Windows.Forms.TrackBar
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents btnDefaults2 As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents tbRotateVertical As System.Windows.Forms.TrackBar
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents tbRotateHorizontal As System.Windows.Forms.TrackBar
    Friend WithEvents btnDefaultSkins As System.Windows.Forms.Button
    Friend WithEvents cbStretch As System.Windows.Forms.CheckBox
    Friend WithEvents cbEnabled As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnLocate As System.Windows.Forms.Button
    Friend WithEvents btnLineGoldenRatio As System.Windows.Forms.Button
    Friend WithEvents btnLineThirds As System.Windows.Forms.Button
    Friend WithEvents btnGetMoreFreeSkinsOnline As System.Windows.Forms.Button
End Class
