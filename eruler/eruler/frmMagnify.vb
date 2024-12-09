Imports System.ComponentModel

Public Class frmMagnify

    'Dim frmghost As Form ' v3.3
    Dim frmghost As Form = New frmGhostAbout

    Private OriginalSize As Integer
    Private hgSuspendTimer2 As Boolean

    Private Sub frmResize_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        hgSuspendTimer2 = gSuspendTimer2
        gSuspendTimer2 = True

        Me.Size = ResizeMe(Me.Size) 'v3.1

        MakeTopMost(Me.Handle.ToInt64)
        OriginalSize = MagnifierFactor

        Me.BackgroundImage = SetImageOpacity(gCurrentSkin.HorizontalBackground, gCurrentSkin.Opacity)

        If RulerIsHorizontal Then
            Me.Location = New Point(gSetLocation.X, gSetLocation.Y)
        Else
            Me.Location = New Point(gSetLocation.X - 34, gSetLocation.Y)
        End If

        'only allow the primary screen to be magnified
        Dim PrimaryScreenBounds As Rectangle = Screen.PrimaryScreen.Bounds

        ' Here are the points of the ruler
        ' A --- B
        ' |     |
        ' C --- D

        Dim A As Point = New Point(frmRuler.Location.X + frmRuler.RulerPanel.Location.X, frmRuler.Location.Y + frmRuler.RulerPanel.Location.Y)
        Dim B As Point = New Point(frmRuler.Location.X + frmRuler.RulerPanel.Location.X + gRulerPanelBounds.Width, frmRuler.Location.Y + frmRuler.RulerPanel.Location.Y)
        Dim C As Point = New Point(frmRuler.Location.X + frmRuler.RulerPanel.Location.X, frmRuler.Location.Y + frmRuler.RulerPanel.Location.Y + gRulerPanelBounds.Height)
        Dim D As Point = New Point(frmRuler.Location.X + frmRuler.RulerPanel.Location.X + gRulerPanelBounds.Width, frmRuler.Location.Y + frmRuler.RulerPanel.Location.Y + gRulerPanelBounds.Height)

        Dim S1 As System.Windows.Forms.Screen = Screen.FromPoint(A)
        Dim S2 As System.Windows.Forms.Screen = Screen.FromPoint(B)
        Dim S3 As System.Windows.Forms.Screen = Screen.FromPoint(C)
        Dim S4 As System.Windows.Forms.Screen = Screen.FromPoint(D)

        Dim A1 As Point = New Point(frmRuler.RulerPanel.Location.X, frmRuler.RulerPanel.Location.Y)
        Dim B1 As Point = New Point(frmRuler.RulerPanel.Location.X + gRulerPanelBounds.Width, frmRuler.RulerPanel.Location.Y)
        Dim C1 As Point = New Point(frmRuler.RulerPanel.Location.X, frmRuler.RulerPanel.Location.Y + gRulerPanelBounds.Height)
        Dim D1 As Point = New Point(frmRuler.RulerPanel.Location.X + gRulerPanelBounds.Width, frmRuler.RulerPanel.Location.Y + gRulerPanelBounds.Height)


        Dim OnPrimaryScreen As Boolean = S1.Primary AndAlso S2.Primary AndAlso S3.Primary AndAlso S4.Primary AndAlso
                                        IsThisPointInBounds(A1) AndAlso IsThisPointInBounds(B1) AndAlso IsThisPointInBounds(C1) AndAlso IsThisPointInBounds(D1)

        If OnPrimaryScreen Then
        Else

            Call MsgBox("The ruler must be fully contained within the primary screen for you to use the magnify feature.", MsgBoxStyle.Information Or MsgBoxStyle.OkOnly, gResourceManager.GetString("0052"))
            Me.DialogResult = DialogResult.None ' keeps form from closing

            Me.Close()

        End If

    End Sub

    Private Sub Option_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbNormal.Click, rb2x.Click, rb3x.Click, rb4x.Click
        MagnifierFactor = CType(sender.tag, Integer)
        Me.Hide()
        Me.Close()
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        e.Handled = True
        If e.KeyCode = Keys.Escape Then
            MagnifierFactor = 1
            Me.Hide()
            Me.Close()
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
            Me.BringToFront()
        End If

        Me.Width = 168
        Me.Height = 102

    End Sub

    Private Sub frm_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move
        If frmghost IsNot Nothing Then frmghost.Location = Me.Location
    End Sub

    Private Sub frm_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If frmghost IsNot Nothing Then frmghost.Dispose()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        'button ok is hidden off the window - it is used as a default when the enter key is pressed
        MagnifierFactor = 1
        Me.Hide()
        Me.Close()
    End Sub

    Private Sub frmMagnify_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        gSuspendTimer2 = hgSuspendTimer2

    End Sub
End Class