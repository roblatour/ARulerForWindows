Imports System.ComponentModel

Public Class FrmNewName

    Private OKPressed As Boolean = False
    Private hgSuspendTimer2 As Boolean

    Private Sub frmSkin_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        hgSuspendTimer2 = gSuspendTimer2
        gSuspendTimer2 = True

        Dim DPI As Integer = SafeNativeMethods.GetDpiForWindow(Me.Handle)

        If DPI = 96 Then
            Me.AutoScaleMode = AutoScaleMode.None
            gUniversaleScale = 1
        Else
            Me.AutoScaleMode = AutoScaleMode.Font
            gUniversaleScale = CSng(DPI / 96)
        End If

        Me.Size = ResizeMe(Me.Size) 'v3.1

        MakeTopMost(Me.Handle.ToInt64)
        RefreshAllSkinNames()

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Dim TestString As String = Me.TextBox1.Text.ToLower.Trim

        If TestString.Contains(gResourceManager.GetString("0044")) OrElse TestString.Contains(gResourceManager.GetString("0045")) Then '  (enabled) or (disabled)
            Call MsgBox(gResourceManager.GetString("0071"), MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, gResourceManager.GetString("0052")) ' Name may not contain '(enabled)' or '(disable)'.
            Me.DialogResult = DialogResult.None ' keeps form from closing
            Exit Sub
        End If

        If IsThisAValidFilename(TestString) Then
        Else
            Call MsgBox(gResourceManager.GetString("0072"), MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, gResourceManager.GetString("0052")) ' "Name contains one or more invalid characters.  Invalid characters are"
            Me.DialogResult = DialogResult.None ' keeps form from closing
            Exit Sub
        End If

        For Each SkinName As String In gAllSkinNames
            If TestString = SkinName.ToLower Then
                Call MsgBox(MsgBox(gResourceManager.GetString("0073"), MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, gResourceManager.GetString("0047"))) ' "Sorry, that name is already in use."
                Me.DialogResult = DialogResult.None ' keeps form from closing
                Exit Sub
            End If
        Next

        gCurrentSkin.Name = Me.TextBox1.Text.Trim
        OKPressed = True
        Me.Close()

    End Sub


    Private Function IsThisAValidFilename(ByVal Filename As String) As Boolean

        Dim ReturnCode As Boolean = True

        If (Filename Is Nothing) OrElse (Filename = "") Then
            ReturnCode = False
        Else
            Dim InvalidFileChars() As Char = System.IO.Path.GetInvalidFileNameChars()
            For x As Integer = 0 To InvalidFileChars.Length - 1
                If Filename.Contains(InvalidFileChars(x).ToString) Then
                    ReturnCode = False
                    Exit For
                End If
            Next
        End If

        Return ReturnCode

    End Function

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub FrmNewName_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If OKPressed Then
        Else
            gCurrentSkin.Name = ""
        End If

    End Sub

    Private Sub FrmNewName_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        gSuspendTimer2 = hgSuspendTimer2
    End Sub
End Class