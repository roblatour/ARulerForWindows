Imports System.ComponentModel

Public Class frmResize

    'Dim frmghost As Form ' v3.3
    Dim frmghost As Form = New frmGhostAbout
    Private hgSuspendTimer2 As Boolean

    Private OriginalSize As Integer

    Private Sub frmResize_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim DPI As Integer = SafeNativeMethods.GetDpiForWindow(Me.Handle)

        If DPI = 96 Then
            Me.AutoScaleMode = AutoScaleMode.None
            gUniversaleScale = 1
        Else
            Me.AutoScaleMode = AutoScaleMode.Font
            gUniversaleScale = CSng(DPI / 96)
        End If

        hgSuspendTimer2 = gSuspendTimer2
        gSuspendTimer2 = True

        ' Me.Size = ResizeMe(Me.Size) 'v3.1

        MakeTopMost(Me.Handle.ToInt64)

        Me.BackgroundImage = SetImageOpacity(gCurrentSkin.HorizontalBackground, gCurrentSkin.Opacity)

        Me.btnOK.BackgroundImage = Me.BackgroundImage
        Me.btnCancel.BackgroundImage = Me.BackgroundImage
        Me.TextBox1.Text = NewSize * gRulerScalingFactorSetByUser

        Me.TextBox1.SelectAll()

        OriginalSize = NewSize * gRulerScalingFactorSetByUser

        If RulerIsHorizontal Then
            Me.Location = gSetLocation
        Else
            Me.Location = New Point(gSetLocation.X - 34, gSetLocation.Y)
        End If

    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Me.TextBox1.Text = Me.TextBox1.Text.Trim

        Try
            NewSize = CSng(TextBox1.Text)
        Catch ex As Exception
            NewSize = 250
        End Try

        NewSize = NewSize / gRulerScalingFactorSetByUser

        If NewSize < 100 Then NewSize = 100

        Me.Close()

    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        NewSize = OriginalSize
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress

        If [Char].IsDigit(e.KeyChar.ToString()) OrElse (e.KeyChar.ToString = ".") Then
        ElseIf e.KeyChar = vbBack Then
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub frm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If gCurrentSkin.Opacity < 1 Then
            If frmghost Is Nothing Then
                frmghost = New frmGhostAbout
                frmghost.ControlBox = False
                frmghost.Size = Me.Size
                frmghost.Show()
                frmghost.Location = Me.Location
            End If
        End If

        Dim TextToUse As String

        'If gShowReadingGuide Then
        '    TextToUse = gResourceManager.GetString("0076") ' Set reading guide length
        'Else
        '    TextToUse = gResourceManager.GetString("0003") ' Set ruler length
        'End If
        TextToUse = gResourceManager.GetString("0003") ' Set length
        'End If

        'Try
        '    If TextToUse.Length < 28 Then
        '        Me.Text = TextToUse
        '    Else
        '        Me.Text = TextToUse
        '    End If
        'Catch ex As Exception
        '    Me.Text = TextToUse
        'End Try

        Me.Text = TextToUse

        Me.Width = 168
        Me.Height = 102

    End Sub

    Private Sub frm_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move
        If frmghost IsNot Nothing Then frmghost.Location = Me.Location
    End Sub

    Private Sub frm_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If frmghost IsNot Nothing Then frmghost.Dispose()
    End Sub

    Private Sub frmResize_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        gSuspendTimer2 = hgSuspendTimer2

    End Sub
End Class