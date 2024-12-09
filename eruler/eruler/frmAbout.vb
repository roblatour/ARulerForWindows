Imports System.ComponentModel

Public Class frmAbout

    Private g As Graphics

    Private Loading As Boolean = True

    Private FirstActivation As Boolean = True

    Dim frmGhostAbout As Form ' v3.3 do not change this line

    Private Sub frmAbout_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        My.Settings.scale = CSng(tbScale.Text.Trim)

        My.Settings.Bump = NumericUpDown1.Value

        My.Settings.Save()

        gRulerScalingFactorSetByUser = My.Settings.scale

        gSuspendTimer2 = hgSuspendTimer2

        Application.DoEvents()

    End Sub

    Private hgSuspendTimer2 As Boolean

    Private DPI As Integer

    Private Sub frmAbout_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        DPI = SafeNativeMethods.GetDpiForWindow(Me.Handle)

        If DPI = 96 Then
            Me.AutoScaleMode = AutoScaleMode.None
            gUniversaleScale = 1
        Else
            Me.AutoScaleMode = AutoScaleMode.Font
            gUniversaleScale = CSng(DPI / 96)
        End If

        hgSuspendTimer2 = gSuspendTimer2
        gSuspendTimer2 = True

        gRefreshRulerAndAboutWindow = False

        Me.Size = ResizeMe(Me.Size) 'v3.1

        'centre window in whatever screen the mouse is currently within
        Dim BoundsOfCurrentScreen As Rectangle = Screen.GetBounds(MousePosition)
        Me.Location = New Point(BoundsOfCurrentScreen.X + BoundsOfCurrentScreen.Width / 2 - Me.Width / 2, BoundsOfCurrentScreen.Y + BoundsOfCurrentScreen.Height / 2 - Me.Height / 2)

        SetBackgroundsOnControls(gCurrentSkin.HorizontalBackground)

        MakeTopMost(Me.Handle.ToInt64)
        cbConfirmExit.Checked = My.Settings.ConfirmOnExit
        cbQuickStartup.Checked = My.Settings.QuickStartup
        cbHelpBalloons.Checked = My.Settings.ShowHelpBalloonsOnRuler
        cbFence.Checked = My.Settings.FenceRulerOnScreen
        cbAutoExpand.Checked = My.Settings.AutoExpand
        NumericUpDown1.Value = My.Settings.Bump
        tbScale.Text = My.Settings.scale

        ResetKeyValues()

        UnderLineFlag()

        HorizontallyCentreLableLocationLocation(Me.lblHeader)
        HorizontallyCentreLableLocationLocation(Me.lblWebSite)
        HorizontallyCentreLableLocationLocation(Me.lblCopyright)

        AboutLabelsandText(lblShiftArrow, lblShiftArrowText)

        Me.ToolTip1.Active = gShowBalloons

        If gPeferedLocationForHelp = gPreferredLocationNotSet Then
        Else
            Me.Location = gPeferedLocationForHelp
            gPeferedLocationForHelp = gPreferredLocationNotSet
        End If

        Loading = False

    End Sub

    Private Sub frm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If FirstActivation Then

            Me.Width = CInt(CSng(gAboutWidth) * gUniversaleScale) 'v3.6.1
            Me.Height = CInt(CSng(gAboutHeight) * gUniversaleScale) 'v3.6.1

            Application.DoEvents()

            'show these in the colour the user chose for the line boxes, likely the most visible
            lblHeader.ForeColor = gCurrentSkin.LineBoxes
            lblWebSite.ForeColor = gCurrentSkin.LineBoxes
            lblCopyright.ForeColor = gCurrentSkin.LineBoxes

        End If

        If gCurrentSkin.Opacity < 1 Then

            If frmGhostAbout Is Nothing Then

                frmGhostAbout = New frmGhostAbout

                If DPI = 96 Then
                    frmGhostAbout.AutoScaleMode = AutoScaleMode.None
                Else
                    frmGhostAbout.AutoScaleMode = AutoScaleMode.Dpi
                End If

                frmGhostAbout.Size = Me.Size

                frmGhostAbout.Show()

                frmGhostAbout.Location = Me.Location

            End If
        End If

        Me.Refresh()
        Application.DoEvents()

        If FirstActivation Then

            FirstActivation = False

            btnOK.BringToFront()
            btnOK.Focus()

        End If

    End Sub

    Private Sub HorizontallyCentreLableLocationLocation(ByRef lbl As Windows.Forms.Label)

        Dim w As Integer = g.MeasureString(lbl.Text, lbl.Font).Width
        Dim NewLocation As Point = New Point(Me.Width / 2 - w / 2, lbl.Location.Y)
        lbl.Location = NewLocation

    End Sub

    Private Sub AboutLabelsandText(ByVal lbl1 As Windows.Forms.Label, ByRef lbl2 As Windows.Forms.Label)

        Dim w As Integer = g.MeasureString(lbl1.Text, lbl1.Font).Width
        Dim NewLocation As Point = New Point(lbl1.Location.X + w + 15, lbl1.Location.Y)
        lbl2.Location = NewLocation

    End Sub

    Private Sub ResetKeyValues()

        Const Quote As String = """"
        Const Collon As String = ":"

        Me.Label_Clear.Text = Quote & gKey_Clear & Quote & Collon
        Me.Label_ShowGolden.Text = Quote & gKey_ShowGoldenRatioLine & Quote & Collon
        Me.Label_ShowLength.Text = Quote & gKey_ShowLength & Quote & Collon
        Me.Label_ShowMidpoint.Text = Quote & gKey_ShowMidPoint & Quote & Collon
        Me.Label_ReverseNumbers.Text = Quote & gKey_ReverseNumbers & Quote & Collon
        Me.Label_SwitchRuleAndReadingGuide.Text = Quote & gKey_SwitchBetweenRulerAndGuide & Quote & Collon
        Me.Label_ShowThirds.Text = Quote & gKey_ShowThirds & Quote & Collon
        Me.Label_PageUpPageDown.Text = gKey_PageUpPageDown & Collon
        Me.Label_Home.Text = gKey_Home & Collon
        Me.Label_Escape.Text = gKey_Escape & Collon

        Me.Label_PageUpPageDown_Text.Location = New Point(Me.Label_PageUpPageDown.Location.X + Me.Label_PageUpPageDown.Width, Me.Label_PageUpPageDown.Location.Y)
        Me.Label_Home_Text.Location = New Point(Me.Label_Home.Location.X + Me.Label_Home.Width, Me.Label_Home.Location.Y)
        Me.Label_Escape_Text.Location = New Point(Me.Label_Escape.Location.X + Me.Label_Escape.Width, Me.Label_Escape.Location.Y)

    End Sub

    Private Sub SetBackgroundsOnControls(ByVal img As Image)

        'Me.BackgroundImage = Nothing 'v3.3 perfer this
        'Me.BackgroundImage.Dispose()

        If gCurrentSkin.Tiled OrElse (gCurrentSkin.Name = "Wood") Then
            Me.BackgroundImageLayout = ImageLayout.Tile
            Application.DoEvents()
        Else
            Me.BackgroundImageLayout = ImageLayout.Stretch
            Application.DoEvents()
        End If

        If gCurrentSkin.Opacity = 1 Then
            Me.BackgroundImage = img
        End If

    End Sub

    Private Sub frm_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move
        If frmGhostAbout IsNot Nothing Then frmGhostAbout.Location = Me.Location
    End Sub

    Private Sub frm_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If frmGhostAbout IsNot Nothing Then frmGhostAbout.Dispose()
    End Sub

    Private Sub lblWebSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblWebSite.Click

        ' StartAProcess(gResourceManager.GetString("0070")) '"Visit the website"

        Const WebPagePart1 As String = "https://github.com/roblatour/ARulerForWindows/blob/main/languages/"

        Dim WebPagePart2 As String

        If My.Settings.Language.StartsWith("ar") Then
            WebPagePart2 = "ar"

        ElseIf My.Settings.Language.StartsWith("de") Then
            WebPagePart2 = "de"

        ElseIf My.Settings.Language.StartsWith("fr") Then
            WebPagePart2 = "fr"

        ElseIf My.Settings.Language.StartsWith("nl") Then
            WebPagePart2 = "nl"

        ElseIf My.Settings.Language.StartsWith("pt") Then
            WebPagePart2 = "pt"

        ElseIf My.Settings.Language.StartsWith("pl") Then
            WebPagePart2 = "pl"

        ElseIf My.Settings.Language.StartsWith("it") Then
            WebPagePart2 = "it"

        ElseIf My.Settings.Language.StartsWith("es") Then
            WebPagePart2 = "es"

        ElseIf My.Settings.Language.StartsWith("sv") Then
            WebPagePart2 = "sv"
        Else
            WebPagePart2 = "en"

        End If

        Const WebPagePart3 As String = "/README.md"

        Dim WebPage As String = WebPagePart1 & WebPagePart2 & WebPagePart3

        StartAProcess(WebPage)

        Application.DoEvents()

        Me.Close()

    End Sub
    Private Sub btnSupportNow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDonate.Click

        Dim WebPage As String = String.Empty

        If My.Settings.Language.StartsWith("ar") Then
            WebPage = "https://buymeacoffee.com/roblatour"

        ElseIf My.Settings.Language.StartsWith("de") Then
            WebPage = "https://buymeacoffee.com/roblatour"

        ElseIf My.Settings.Language.StartsWith("fr") Then
            WebPage = "https://buymeacoffee.com/roblatour"

        ElseIf My.Settings.Language.StartsWith("nl") Then
            WebPage = "https://buymeacoffee.com/roblatour"

        ElseIf My.Settings.Language.StartsWith("pt") Then
            WebPage = "https://buymeacoffee.com/roblatour"

        ElseIf My.Settings.Language.StartsWith("pl") Then
            WebPage = "https://buymeacoffee.com/roblatour"

        ElseIf My.Settings.Language.StartsWith("it") Then
            WebPage = "https://buymeacoffee.com/roblatour"

        ElseIf My.Settings.Language.StartsWith("es") Then
            WebPage = "https://buymeacoffee.com/roblatour"

        ElseIf My.Settings.Language.StartsWith("sv") Then
            WebPage = "https://buymeacoffee.com/roblatour"

        Else
            WebPage = "https://buymeacoffee.com/roblatour"

        End If

        StartAProcess(WebPage)
        Application.DoEvents()

        Me.Close()

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbConfirmExit.CheckedChanged
        If Loading Then Exit Sub
        My.Settings.ConfirmOnExit = cbConfirmExit.Checked
        My.Settings.Save()
        Application.DoEvents()
    End Sub
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbHelpBalloons.CheckedChanged

        If Loading Then Exit Sub

        gShowBalloons = cbHelpBalloons.Checked
        My.Settings.ShowHelpBalloonsOnRuler = cbHelpBalloons.Checked
        My.Settings.Save()

        Me.ToolTip1.Active = gShowBalloons

        Application.DoEvents()

    End Sub
    Private Sub cbQuickStartup_CheckedChanged(sender As Object, e As EventArgs) Handles cbQuickStartup.CheckedChanged

        If Loading Then Exit Sub

        gQuickStartup = cbQuickStartup.Checked
        My.Settings.QuickStartup = cbQuickStartup.Checked
        My.Settings.Save()

        If gQuickStartup Then
            ShortCut("Add")
        Else
            ShortCut("Delete")
        End If

        Application.DoEvents()

    End Sub

    Private Sub CbAutoExpand_CheckedChanged(sender As Object, e As EventArgs) Handles cbAutoExpand.CheckedChanged
        If Loading Then Exit Sub
        gAutoExpand = cbAutoExpand.Checked
        My.Settings.AutoExpand = cbAutoExpand.Checked
        My.Settings.Save()
        Application.DoEvents()
    End Sub


    Private Sub cbFence_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbFence.CheckedChanged
        If Loading Then Exit Sub
        gFenceRulerOnScreen = cbFence.Checked
        My.Settings.FenceRulerOnScreen = cbFence.Checked
        My.Settings.Save()
        Application.DoEvents()
    End Sub
    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        If Loading Then Exit Sub
        gBump = NumericUpDown1.Value
        'move to forms closing to prevent error being thrown

        'My.Settings.Bump = NumericUpDown1.Value
        'My.Settings.Save()
        'Application.DoEvents()
    End Sub

    Private Sub frmAbout_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        Dim x As Integer = 5 * gUniversaleScale
        Dim i As Integer = 20 * gUniversaleScale
        Dim y As Integer = Me.Label17.Location.Y * gUniversaleScale + i

        g = SymbolPanel.CreateGraphics

        'added in v2.4.4
        Me.Font = Label21.Font
        Me.cbConfirmExit.Font = Label21.Font
        Me.cbFence.Font = Label21.Font
        Me.cbHelpBalloons.Font = Label21.Font
        'end change v2.4.4

        'Temporarily set colour of control boxes to black while they are drawn on this screen
        Dim holdBoxesColour As Color = gCurrentSkin.LineBoxes
        gCurrentSkin.LineBoxes = Color.Black

        DrawASquareInABox(g, New Rectangle(x, y, 10, 10)) : y += i
        DrawANumberSignBox(g, New Rectangle(x, y, 10, 10)) : y += i
        DrawASlashInABox(g, New Rectangle(x, y, 10, 10)) : y += i
        DrawAPlusInABox(g, New Rectangle(x, y, 10, 10)) : y += i
        DrawTicksInABox(g, New Rectangle(x, y, 10, 10), True, True) : y += 2 * i
        DrawACircleInABox(g, New Rectangle(x, y, 10, 10)) : y += i
        DrawAQuestionMarkInABox(g, New Rectangle(x, y, 10, 10)) : y += i
        DrawAMinusInABox(g, New Rectangle(x, y, 10, 10)) : y += i
        DrawAnXInABox(g, New Rectangle(x, y, 10, 10))

        gCurrentSkin.LineBoxes = holdBoxesColour

    End Sub

    Private Sub btnOK_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Dim NewScale As Single

        Try
            NewScale = CSng(tbScale.Text.Trim)
        Catch ex As Exception
            Beep()
            Dim Title As String = gResourceManager.GetString("0052")
            Call MessageBox.Show(lblScale.Text & " <> " & tbScale.Text.Trim, Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.DialogResult = DialogResult.None ' keeps form from closing
            Exit Sub
        End Try


        If NewScale = CSng(0.0625) Then
        Else
            If NewScale > CSng(5) Then
                Beep()
                Dim Title As String = gResourceManager.GetString("0052")
                Call MessageBox.Show(lblScale.Text & " > 5.0", Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.DialogResult = DialogResult.None ' keeps form from closing
                Exit Sub

            ElseIf NewScale = CSng(0) Then
                Beep()
                Dim Title As String = gResourceManager.GetString("0052")
                Call MessageBox.Show(lblScale.Text & " <> 0", Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.DialogResult = DialogResult.None ' keeps form from closing
                Exit Sub

            ElseIf NewScale < CSng(0.5) Then
                Beep()
                Dim Title As String = gResourceManager.GetString("0052")
                Call MessageBox.Show(lblScale.Text & " < 0.5", Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Me.DialogResult = DialogResult.None ' keeps form from closing
                Exit Sub
            End If
        End If

        Me.Close()

    End Sub

    Private Sub btnSkins_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSkins.Click

        ' Hide this form and get rid of the ghost
        '(The ghost will return when this form as soon as this form is shown again)

        Me.Hide()

        If frmGhostAbout IsNot Nothing Then
            frmGhostAbout.Dispose()
            frmGhostAbout = Nothing
        End If

        Dim frmSkins As Form = New frmSkins
        frmSkins.ShowDialog()
        frmSkins.Dispose()

        Me.TransparencyKey = gCurrentSkin.TransparencyColour
        MyTransparentColour = gCurrentSkin.TransparencyColour

        RefreshAllSkinNames()
        While Not gCurrentSkin.SkinEnabled
            AdvanceToNextSkin()
        End While

        Me.Close()

    End Sub

    Private Sub pbLicense_Click(sender As Object, e As EventArgs) Handles pbLicense.Click

        StartAProcess(gLicenseWebSite)
        Me.Close()

    End Sub

    Private Sub Flag_Click(sender As Object, e As EventArgs) Handles pbEN_UK.Click, pbEN_US.Click, pbAR.Click, pbDE.Click, pbES.Click, pbFR.Click, pbIT.Click, pbNL.Click, pbPL.Click, pbPT.Click, pbSV.Click

        My.Settings.Language = sender.tag
        My.Settings.Save()

        SetLanguage()

        gPeferedLocationForHelp = Me.Location
        gReloadHelp = True

        Me.Close()

    End Sub

    Private Sub UnderLineFlag()

        Dim wLangCode As String = My.Settings.Language.ToUpper

        If wLangCode = "EN-US" Then
            Me.lbDotUS.Visible = True

        ElseIf wLangCode.StartsWith("AR") Then
            Me.lbDotAR.Visible = True

        ElseIf wLangCode.StartsWith("DE") Then
            Me.lbDotDE.Visible = True
            Me.pbDE.BorderStyle = BorderStyle.FixedSingle

        ElseIf wLangCode.StartsWith("ES") Then
            Me.lbDotES.Visible = True

        ElseIf wLangCode.StartsWith("FR") Then
            Me.lbDotFR.Visible = True

        ElseIf wLangCode.StartsWith("IT") Then
            Me.lbDotIT.Visible = True

        ElseIf wLangCode.StartsWith("NL") Then
            Me.lbDotNL.Visible = True

        ElseIf wLangCode.StartsWith("PL") Then
            Me.lbDotPL.Visible = True

        ElseIf wLangCode.StartsWith("PT") Then
            Me.lbDotPT.Visible = True

        ElseIf wLangCode.StartsWith("SV") Then
            Me.lbDotSV.Visible = True

        Else
            Me.lbDotUK.Visible = True

        End If

    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbScale.KeyPress

        If [Char].IsDigit(e.KeyChar.ToString()) OrElse (e.KeyChar.ToString = ".") Then
        ElseIf e.KeyChar = vbBack Then
        Else
            e.Handled = True
        End If

    End Sub

End Class