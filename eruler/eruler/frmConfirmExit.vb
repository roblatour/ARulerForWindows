
Imports System.ComponentModel

Public Class frmConfirmExit

    'Dim frmghost As Form ' v3.3
    Dim frmghost As Form = New frmGhostAbout

    Private hgSuspendTimer2 As Boolean

    Private Sub frmConfirmExit_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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

        MakeTopMost(Me.Handle.ToInt64)

        Me.Size = ResizeMe(Me.Size) 'v3.1

        Me.BackgroundImage = SetImageOpacity(gCurrentSkin.HorizontalBackground, gCurrentSkin.Opacity)

        Me.btnDoNotExit.BackgroundImage = Me.BackgroundImage
        Me.btnExit.BackgroundImage = Me.BackgroundImage

        If RulerIsHorizontal Then
            Me.Location = New Point(gSetLocation.X + 5, gSetLocation.Y + 8)
        Else
            Me.Location = New Point(gSetLocation.X - 53, gSetLocation.Y + 5)
            Me.btnDoNotExit.Select()
        End If

        'added in v4.2.4
        btnDoNotExit.Font = Me.Font
        btnExit.Font = Me.Font
        'end of change v2.4.2

    End Sub

    Private Sub button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click, btnDoNotExit.Click

        gOKToExitNow = CType(sender.tag, Boolean)

        If sender.tag = "TRUE" Then
            frmRuler.CloseOnDemand = True
            frmRuler.CloseOnlyOnce = False
            Me.Close()
        Else
            gOKToExitNow = False
        End If

    End Sub

    Private Sub frm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If gCurrentSkin.Opacity < 1 Then
            If frmghost Is Nothing Then
                frmghost = New frmGhostAbout
                frmghost.Size = Me.Size
                frmghost.Show()
                frmghost.Location = Me.Location
            End If
        End If

        Me.Width = 206
        Me.Height = 83

    End Sub

    Private Sub frm_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move
        If frmghost IsNot Nothing Then frmghost.Location = Me.Location
    End Sub

    Private Sub frm_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If frmghost IsNot Nothing Then frmghost.Dispose()
    End Sub

    Private Sub frmConfirmExit_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        gSuspendTimer2 = hgSuspendTimer2
    End Sub
End Class