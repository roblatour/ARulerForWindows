Imports System.ComponentModel

Public Class frmSkins

    Private Loading As Boolean = True

    Private FirstActivation As Boolean = True

    Private WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog = New OpenFileDialog

    Private ColourDialog As ColorDialog = New ColorDialog

    Private OkButtonPressed As Boolean = False

    Dim HoldSkin As SkinStructure

    Private CurrentSelectedSkin As String

    Private frmGhost As Form

    Private hgSuspendTimer2 As Boolean

    Const gSkinWidth As Integer = 489
    Const gSinkHeight As Integer = 489

    Private Sub frmSkin_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

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

            Application.DoEvents()

            'Me.Size = ResizeMe(Me.Size) 'v3.1

            Me.VerticalRuler.Size = New Size(Me.VerticalRuler.Size.Width, Me.VerticalRuler.Size.Height)
            Me.HorizontalRuler.Size = New Size(Me.HorizontalRuler.Size.Width, Me.HorizontalRuler.Size.Height)

            'centre window in whatever screen the mouse is currently within
            Dim BoundsOfCurrentScreen As Rectangle = Screen.GetBounds(MousePosition)
            Me.Location = New Point(BoundsOfCurrentScreen.X + BoundsOfCurrentScreen.Width / 2 - Me.Width / 2, BoundsOfCurrentScreen.Y + BoundsOfCurrentScreen.Height / 2 - Me.Height / 2)

            MakeTopMost(Me.Handle.ToInt64)

            gRefreshRulerAndAboutWindow = True

            btnSave.Enabled = False

            btnSave.Visible = False

            Loading = False

        Catch ex As Exception

            MsgBox("frmSkin_Load " & vbCrLf & ex.Message.ToString)

        End Try

    End Sub

    Private Sub frm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Try

            If FirstActivation Then

                FirstActivation = False

                Application.DoEvents()

                Me.ToolTip1.Active = gShowBalloons

                MyTransparentColour = gCurrentSkin.TransparencyColour

            End If

            If frmGhost Is Nothing Then

                frmGhost = New frmGhostSkin
                frmGhost.Size = Me.Size
                frmGhost.Show()
                frmGhost.Location = Me.Location
                Application.DoEvents()

                RefreshAllSkinNames()

                If gAllSkinNames Is Nothing Then
                Else

                    Dim AllTranslatedSkinNames As String() = gAllSkinNames
                    For x As Integer = 0 To AllTranslatedSkinNames.Length - 1
                        AllTranslatedSkinNames(x) = TranslateASkinNameFromEnglishtToAnotherLanguage(AllTranslatedSkinNames(x))
                    Next

                    Me.cbSkinNames.Items.AddRange(AllTranslatedSkinNames)

                    For x As Integer = 0 To cbSkinNames.Items.Count - 1

                        If GetEnabledSetting(cbSkinNames.Items(x)) Then
                            cbSkinNames.Items(x) &= " " & gResourceManager.GetString("0044") '(enabled)
                        Else
                            cbSkinNames.Items(x) &= " " & gResourceManager.GetString("0045") '(disabled)
                        End If

                    Next

                    LoadFirstSkin()

                    HoldSkin = gCurrentSkin

                    btnRemove.Enabled = (cbSkinNames.Items.Count > 1)

                End If

            End If

            MakeTopMost(Me.Handle.ToInt64)

        Catch ex As Exception

            MsgBox("frmSkin_Load " & vbCrLf & ex.Message.ToString)

        End Try

    End Sub

    Private Sub frmSkin_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If OkButtonPressed Then
        Else
            'window is closing as user clicked the red 'x' in the top right hand courner of the window

            If btnSave.Enabled Then
                If ConfirmIfCurrentSkinShouldBeSaved() Then
                Else
                    gCurrentSkin = HoldSkin
                End If
            End If
        End If

        My.Settings.CurrentSkinName = gCurrentSkin.Name
        My.Settings.Save()
        Application.DoEvents()

    End Sub

    Private Sub frm_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move

        On Error Resume Next
        If frmGhost IsNot Nothing Then
            frmGhost.Location = Me.Location
            Me.Refresh()
        End If


    End Sub

    Private Sub frm_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

        If frmGhost IsNot Nothing Then frmGhost.Dispose()

    End Sub

    Private Sub btnLineBoxes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLineBoxes.Click, btnLineLength.Click, btnLineMeasuring.Click, btnLineMidpoint.Click, btnNumbers.Click, btnLineTicksAndSides.Click, btnLineGoldenRatio.Click, btnLineThirds.Click

        ColourDialog.FullOpen = True
        ColourDialog.AnyColor = True
        ColourDialog.ShowHelp = False

        ColourDialog.CustomColors = LoadMyCustomColours()

        Dim RetCode As DialogResult = ColourDialog.ShowDialog()

        If RetCode = Windows.Forms.DialogResult.OK Then

            Dim AChangeWasMade As Boolean = False

            Select Case sender.name

                Case Is = "btnLineBoxes"
                    If gCurrentSkin.LineBoxes = ColourDialog.Color Then
                    Else
                        AChangeWasMade = True
                        gCurrentSkin.LineBoxes = ColourDialog.Color
                    End If

                Case Is = "btnLineTicksAndSides"
                    If gCurrentSkin.LineTicksAndSides = ColourDialog.Color Then
                    Else
                        AChangeWasMade = True
                        gCurrentSkin.LineTicksAndSides = ColourDialog.Color
                    End If

                Case Is = "btnLineLength"
                    If gCurrentSkin.LineLength = ColourDialog.Color Then
                    Else
                        AChangeWasMade = True
                        gCurrentSkin.LineLength = ColourDialog.Color
                    End If


                Case Is = "btnLineMeasuring"
                    If gCurrentSkin.LineMeasuring = ColourDialog.Color Then
                    Else
                        AChangeWasMade = True
                        gCurrentSkin.LineMeasuring = ColourDialog.Color
                    End If


                Case Is = "btnLineMidpoint"
                    If gCurrentSkin.LineMidpoint = ColourDialog.Color Then
                    Else
                        AChangeWasMade = True
                        gCurrentSkin.LineMidpoint = ColourDialog.Color
                    End If

                Case Is = "btnLineThirds"
                    If gCurrentSkin.LineThirds = ColourDialog.Color Then
                    Else
                        AChangeWasMade = True
                        gCurrentSkin.LineThirds = ColourDialog.Color
                    End If

                Case Is = "btnLineGoldenRatio"
                    If gCurrentSkin.LineGoldenRatio = ColourDialog.Color Then
                    Else
                        AChangeWasMade = True
                        gCurrentSkin.LineGoldenRatio = ColourDialog.Color
                    End If

                Case Is = "btnNumbers"
                    If gCurrentSkin.Numbers = ColourDialog.Color Then
                    Else
                        AChangeWasMade = True
                        gCurrentSkin.Numbers = ColourDialog.Color
                    End If

            End Select

            If AChangeWasMade Then
                btnSave.Enabled = True
                UpdateRulerBackgroundsInGhostWindow()
            End If


        End If

        SaveMyCustomColours(ColourDialog.CustomColors)

        Me.Refresh()

    End Sub

    Private Function LoadMyCustomColours() As Integer()

        'the 16 custom colours are stored in mysettings.customcolors as a single string
        'the string's format is: "number, number, number ..." 
        'this function translates the string into an integer array (0..15)
        'the default value of each custom colour is white

        Static Dim cWhite As Integer = &HFFFFFF ' the colour code value for white is Hex FFFFFF

        Dim iColourTable() As Integer = {cWhite, cWhite, cWhite, cWhite, cWhite, cWhite, cWhite, cWhite, cWhite, cWhite, cWhite, cWhite, cWhite, cWhite, cWhite, cWhite}
        Dim sColourTable() As String = My.Settings.CustomColours.Split(",")

        Try

            For x As Integer = 0 To sColourTable.Length - 1

                If x > 15 Then Exit For

                If sColourTable(x).Length > 0 Then
                    iColourTable(x) = CType(sColourTable(x), Integer)
                End If

            Next

        Catch ex As Exception
            For x As Integer = 0 To 15 : iColourTable(x) = cWhite : Next
        End Try


        Return iColourTable

    End Function

    Private Sub SaveMyCustomColours(ByVal ColourTable() As Integer)

        Dim MyCustomColours As String = CType(ColourTable(0), String)
        For x As Integer = 1 To 15
            MyCustomColours &= " , " & CType(ColourTable(x), String)
        Next

        My.Settings.CustomColours = MyCustomColours
        My.Settings.Save()

    End Sub

    Private Sub btnBackground_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackground.Click

        Try

            If gCurrentSkin.BackgroundImageFileName = "none" Then

            Else

                Dim msg As String = String.Empty
                If gCurrentSkin.BackgroundImageFileName = "embedded" Then
                    msg = gResourceManager.GetString("0056")  ' "Do you want to replace the default image for the this skin?"
                Else
                    msg = gResourceManager.GetString("0057") ' "Do you want to replace the image currently in use for the this skin?"
                End If

                If MsgBox(msg, MsgBoxStyle.YesNo Or MsgBoxStyle.Question, gResourceManager.GetString("0043")) = MsgBoxResult.No Then
                    Exit Try
                End If

            End If

            Application.DoEvents()

            With OpenFileDialog1

                .Title = gResourceManager.GetString("0053") '"A Ruler for Windows - Image Locate"
                ' .Filter = "Image Files (*.bmp;*.gif;*.jpg;*.png;*.tif)|*.bmp;*.gif;*.jpg;*.png;*.tif|All files (*.*)|*.*"
                .Filter = gResourceManager.GetString("0054") & " (*.bmp;*.gif;*.jpg;*.png;*.tif)|*.bmp;*.gif;*.jpg;*.png;*.tif|" & gResourceManager.GetString("0055") & " (*.*)|*.*"
                .InitialDirectory = My.Settings.LastDirectoryLocation
                .SupportMultiDottedExtensions = True
                .AutoUpgradeEnabled = True
                .DereferenceLinks = False

PleaseTryAgain: If .ShowDialog() = DialogResult.OK Then

                    'there is a bug in the opendialog control that allows url's be selected
                    'the following, in conjunction ".DereferenceLinks = False" (as coded above), gets around it
                    Dim ext As String = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
                    If (String.Compare(ext, ".url", True) = 0) Then

                        ' 0051 = "That was not an image file, please try again.
                        ' 0052 = A Ruler for Windows - Warning
                        Call MsgBox(gResourceManager.GetString("0051"), MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, gResourceManager.GetString("0052"))

                        Me.DialogResult = DialogResult.None ' keeps form from closing
                        GoTo PleaseTryAgain
                    End If

                    btnSave.Enabled = True

                    tbTransparency.Value = 100
                    tbRotateVertical.Value = 1
                    tbRotateHorizontal.Value = 1

                    Dim myPic As New Bitmap(OpenFileDialog1.FileName)

                    Dim h, w As Integer

                    h = myPic.Height
                    w = myPic.Width

                    'Make the shortest side the side to be compressed
                    If h > w Then
                        myPic.RotateFlip(RotateFlipType.Rotate90FlipNone)
                        h = myPic.Height
                        w = myPic.Width
                    End If


                    'v 3.6.1
                    gCurrentBaseBitMap = New Bitmap(myPic)
                    'If h = 100 Then

                    '    gCurrentBaseBitMap = New Bitmap(myPic)

                    'Else

                    '    Dim ratio As Double = 100 / h
                    '    Dim NewSize As Double = w * ratio
                    '    w = CType(NewSize, Integer)
                    '    Dim myThumb As System.Drawing.Image = myPic.GetThumbnailImage(w, 100, Nothing, Nothing)
                    '    myPic = Nothing
                    '    gCurrentBaseBitMap = New Bitmap(myThumb)

                    'End If

                    myPic = Nothing

                    SetRotatedHorizontalBitmap()
                    SetRotatedVeritcalBitmap()

                    gCurrentSkin.TransparencyColour = FindAnUnusedColour(gCurrentBaseBitMap)

                    MyTransparentColour = gCurrentSkin.TransparencyColour

                    UpdateRulerBackgroundsInGhostWindow()

                    gCurrentSkin.BackgroundImageFileName = OpenFileDialog1.FileName

                    My.Settings.LastDirectoryLocation = FindPath(.FileName)
                    My.Settings.Save()

                    RefreshAllSkinNames()

                End If

            End With

        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        Call ConfirmIfCurrentSkinShouldBeSaved()
        btnSave.Enabled = False

        Dim HoldCurrentName As String = gCurrentSkin.Name
        HoldSkin = gCurrentSkin

        Dim frmNewName As Form = New FrmNewName
        frmNewName.ShowDialog()
        frmNewName.Dispose()

        If gCurrentSkin.Name.Length > 0 Then

            cbSkinNames.Text = gCurrentSkin.Name
            cbSkinNames.Items.Add(gCurrentSkin.Name & " " & gResourceManager.GetString("0044")) ' (enabled)
            For x As Integer = 0 To cbSkinNames.Items.Count - 1
                If FilterSkinName(cbSkinNames.Items(x)) = gCurrentSkin.Name Then
                    Loading = True
                    cbSkinNames.SelectedIndex = x
                    Loading = False
                    Exit For
                End If
            Next

            SetDefaultSkin()

            SaveCurrentSkin(gCurrentSkin.Name)
            UpdateSkinNameDropDownBox()

            RefreshAllSkinNames()
            SetNextAndPreviousEnabledSkinNames()
            HoldSkin = gCurrentSkin

            btnSave.Enabled = False

        Else

            gCurrentSkin.Name = HoldCurrentName

        End If

        Me.Refresh()

    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click

        On Error Resume Next

        '0042 = "Remove (delete):"
        '0042 = "A Ruler for Windows - Confirm

        If MsgBox(gResourceManager.GetString("0042") & " " & TranslateASkinNameFromEnglishtToAnotherLanguage(gCurrentSkin.Name), MsgBoxStyle.YesNo Or MsgBoxStyle.Question, gResourceManager.GetString("0043")) = MsgBoxResult.Yes Then

            btnSave.Enabled = False

            'find the postion of the current skin name
            Dim SkinNamePositionInDropDownList As Integer
            For SkinNamePositionInDropDownList = 0 To cbSkinNames.Items.Count - 1
                If TranslateASkinNameFromAnotherLanguageToEnglish(FilterSkinName(cbSkinNames.Items(SkinNamePositionInDropDownList))) = gCurrentSkin.Name Then
                    cbSkinNames.SelectedIndex = SkinNamePositionInDropDownList
                    Exit For
                End If
            Next

            RemoveSkin(gCurrentSkin.Name)
            Application.DoEvents()

            'use the skin name in the same relative location in the list as the old name before it was deleted
            '(effectively moving it down one)
            'if it was the last one in the list, back it up by one

            If SkinNamePositionInDropDownList > cbSkinNames.Items.Count - 1 Then
                SkinNamePositionInDropDownList -= 1
            End If

            btnSave.Enabled = False

            cbSkinNames.Text = cbSkinNames.Items(SkinNamePositionInDropDownList)

            cbSkinNames.Refresh()
            gCurrentSkin.Name = FilterSkinName(cbSkinNames.Items(SkinNamePositionInDropDownList))
            Call LoadSkin(gCurrentSkin.Name)

            HoldSkin = gCurrentSkin
            MakeWindowMatchCurrentSkin()

            RefreshAllSkinNames()
            SetNextAndPreviousEnabledSkinNames()

        End If

        btnRemove.Enabled = (cbSkinNames.Items.Count > 1)

    End Sub

    Private Sub MakeWindowMatchCurrentSkin()

        Loading = True

        Me.cbEnabled.Checked = gCurrentSkin.SkinEnabled
        Me.cbStretch.Checked = gCurrentSkin.Tiled

        Me.tbRotateHorizontal.Value = gCurrentSkin.HorizontalRotation
        Me.tbRotateVertical.Value = gCurrentSkin.VerticalRotation

        Me.tbTransparency.Value = gCurrentSkin.Opacity * 100

        Loading = False

        UpdateRulerBackgroundsInGhostWindow()

    End Sub

    Private Sub tbTransparency_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbTransparency.ValueChanged

        If Loading Then Exit Sub

        btnSave.Enabled = True

        Static Dim SkipUpdatesUnder10FromNowOn As Boolean = False

        If (tbTransparency.Value < 10) Then

            If SkipUpdatesUnder10FromNowOn Then Exit Sub

            SkipUpdatesUnder10FromNowOn = True

            gCurrentSkin.Opacity = 1
            gCurrentSkin.TransparencyColour = FindAnUnusedColour(gCurrentSkin.HorizontalBackground)
            UpdateRulerBackgroundsInGhostWindow()
            Application.DoEvents()

            If SetSkinTransparencyColour() Then
            Else
                gCurrentSkin.TransparencyColour = FindAnUnusedColour(gCurrentSkin.HorizontalBackground)
            End If

            tbTransparency.Value = 100

        Else

            SkipUpdatesUnder10FromNowOn = False
            gCurrentSkin.Opacity = CType(tbTransparency.Value / 100, Single)

        End If

        Application.DoEvents()
        UpdateRulerBackgroundsInGhostWindow()

    End Sub

    Private WaitForMouseClick As Boolean
    Private ColourUnderMousePointer As Color

    Private Sub VerticalRuler_MouseLeave(sender As Object, e As EventArgs) Handles VerticalRuler.MouseLeave, HorizontalRuler.MouseLeave

        If WaitForMouseClick Then
            Cursor = Cursors.No
        End If

    End Sub
    Private Sub VerticalRuler_Mouseenter(sender As Object, e As EventArgs) Handles VerticalRuler.MouseEnter, HorizontalRuler.MouseEnter

        If WaitForMouseClick Then
            Cursor = Cursors.Cross
        End If

    End Sub

    Private Function SetSkinTransparencyColour() As Boolean

        Dim ReturnValue As Boolean

        'A Ruler for Windows - Select a transparency colour
        Dim Title As String = gResourceManager.GetString("0067")
        Title = Title.Remove(Title.IndexOf("-") + 2)
        Title &= gResourceManager.GetString("0078")
        ' Dim Msg As String = gResourceManager.GetString("0079").Replace(" | ", vbCrLf & vbCrLf) message too long for resource manager
        Dim Msg As String = GetLongMessage.Replace(" | ", vbCrLf & vbCrLf)

        If MessageBox.Show(Msg, Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = MsgBoxResult.Ok Then

            ''Hide this window to make it clear where the user is to click
            'frmGhost.Opacity = 0
            'Me.Opacity = 0
            'Application.DoEvents()

            gSameLocationAsSkinWindow = Me.Location
            Dim frmPickAColour As Form = New frmPickAColour  'sets gCurrentSkin.TransparencyColour
            frmPickAColour.ShowDialog()

            'frmGhost.Opacity = 1
            'Me.Opacity = 1

            ReturnValue = True

        Else

            ReturnValue = False

        End If

        Return ReturnValue

    End Function


    Private Function GetLongMessage() As String

        Dim wLangCode As String = My.Settings.Language.ToUpper

        Select Case wLangCode

            Case = "EN-US"
                Return "Making the entire background fully transparent would render the ruler/reading guide effectively invisible. | However, you can make portions of your ruler / reading guide skin fully transparent by selecting a color on your skin's background to be used as a transparency mask.  | If you would like to do this, just click 'OK' and new window will appear with your skin's background in it.   After that, just click on any color on that image to make that color the transparency mask.  | If you don't want any part of your background image to be fully transparent just click 'Cancel'."

            Case = "DE"
                Return "Den gesamten Hintergrund vollständig transparent zu machen, würde das Lineal/die Lesehilfe effektiv unsichtbar machen. | Sie können jedoch Teile Ihrer Lineal-/Lesehilfe-Skin vollständig transparent machen, indem Sie eine Farbe auf dem Hintergrund Ihrer Skin auswählen, die als Transparenzmaske verwendet werden soll. | Wenn Sie dies tun möchten, klicken Sie einfach auf 'OK' und ein neues Fenster mit dem Hintergrund Ihres Skins wird angezeigt. Klicken Sie danach einfach auf eine beliebige Farbe In diesem Bild, um diese Farbe zur Transparenzmaske zu machen. | Wenn Sie nicht möchten, dass ein Teil Ihres Hintergrundbilds vollständig transparent ist, klicken Sie einfach auf „Abbrechen""."

            Case = "ES"
                Return "Hacer que todo el fondo sea completamente transparente haría que la regla/guía de lectura fuera efectivamente invisible. | Sin embargo, puede hacer que partes de la máscara de su regla/guía de lectura sean completamente transparentes seleccionando un color en el fondo de su máscara para usar como máscara de transparencia. | Si desea hacer esto, simplemente haga clic en 'Aceptar' y aparecerá una nueva ventana con el fondo de su máscara. Después de eso, simplemente haga clic en cualquier color de esa imagen para convertir ese color en la máscara de transparencia. | Si no desea que ninguna parte de su imagen de fondo sea completamente transparente, simplemente haga clic en 'Cancelar'."

            Case = "FR"
                Return " Rendre l'arrière-plan entièrement transparent rendrait la règle/le guide de lecture effectivement invisible. | Cependant, vous pouvez rendre des parties de l'habillage de votre règle/guide de lecture entièrement transparentes en sélectionnant une couleur sur l'arrière-plan de votre habillage à utiliser comme masque de transparence. | Si vous souhaitez le faire, cliquez simplement sur 'OK' et une nouvelle fenêtre apparaîtra avec l'arrière-plan de votre skin. Après cela, cliquez simplement sur n'importe quelle couleur de cette image pour faire de cette couleur le masque de transparence. | Si vous ne souhaitez pas qu'une partie de votre image d'arrière-plan soit entièrement transparente, cliquez simplement sur ""Annuler""."

            Case = "IT"
                Return "Rendere l'intero sfondo completamente trasparente renderebbe il righello/guida di lettura effettivamente invisibile. | Tuttavia, puoi rendere completamente trasparenti parti della pelle del righello/guida alla lettura selezionando un colore sullo sfondo della pelle da utilizzare come maschera di trasparenza. | Se desideri farlo, fai semplicemente clic su 'OK' e apparirà una nuova finestra con lo sfondo della tua pelle. Dopodiché, fai clic su qualsiasi colore su quell'immagine per rendere quel colore la maschera di trasparenza. | Se non vuoi che nessuna parte dell'immagine di sfondo sia completamente trasparente, fai clic su 'Annulla'."

            Case = "NL"
                Return "Als je de hele stijl volledig transparant maakt, wordt de meetlat/leesregel in feite onzichtbaar. | Je kunt echter delen van de meetlat- / leesregelstijl volledig transparant maken door een kleur van de stijlachtergrond te selecteren en te gebruiken als transparant masker.  | Als je dit zou willen doen, klik je op 'OK' en een nieuw venster zal zich openen met de achtergrondstijl erin. Daarna klik je op een willekeurige kleur in die afbeelding om met die kleur het transparantie masker te maken.  | Als je niet wilt dat een deel van uw stijlafbeelding volledig transparant is, klik dan op 'Cancel'."

            Case = "PL"
                Return "Pełne przeźroczystość całego tła sprawi, że linijka/przewodnik do czytania będzie skutecznie niewidoczny. | Możesz jednak sprawić, że fragmenty skóry linijki / przewodnika czytania będą w pełni przezroczyste, wybierając kolor tła skóry, który będzie używany jako maska przezroczystości. | Jeśli chcesz to zrobić, po prostu kliknij „OK"", a pojawi się nowe okno z tłem Twojej skóry. Następnie wystarczy kliknąć dowolny kolor na tym obrazie, aby ten kolor stał się maską przezroczystości. | Jeśli nie chcesz, aby jakakolwiek część obrazu tła była w pełni przezroczysta, po prostu kliknij ""Anuluj""."

            Case = "PT"
                Return "Tornar todo o plano de fundo totalmente transparente tornaria a régua/guia de leitura efetivamente invisível. | No entanto, você pode tornar partes de sua régua/guia de leitura totalmente transparentes selecionando uma cor no plano de fundo da sua pele para ser usada como máscara de transparência. | Se você quiser fazer isso, basta clicar em 'OK' e uma nova janela aparecerá com o plano de fundo da sua skin. Depois disso, basta clicar em qualquer cor dessa imagem para fazer dessa cor a máscara de transparência. | Se você não quiser que nenhuma parte de sua imagem de fundo seja totalmente transparente, basta clicar em 'Cancelar'.
"
            Case = "SV"
                Return " Att göra hela bakgrunden helt transparent skulle göra linjalen/läsguiden osynlig. | Du kan dock göra delar av din linjal/läsguides hud helt genomskinlig genom att välja en färg på din huds bakgrund som ska användas som en genomskinlig mask. | Om du vill göra detta klickar du bara på ""OK"" så kommer ett nytt fönster upp med din huds bakgrund. Efter det klickar du bara på valfri färg på bilden för att göra den färgen till genomskinlighetsmasken. | Om du inte vill att någon del av din bakgrundsbild ska vara helt genomskinlig klickar du bara på ""Avbryt""."

            Case Else
                Return "Making the entire background fully transparent would render the ruler/reading guide effectively invisible. | However, you can make portions of your ruler / reading guide skin fully transparent by selecting a colour on your skin's background to be used as a transparency mask.  | If you would like to do this, just click 'OK' and new window will appear with your skin's background in it.   After that, just click on any colour on that image to make that colour the transparency mask.  | If you don't want any part of your background image to be fully transparent just click 'Cancel'."

        End Select

    End Function

    Private Sub UpdateRulerBackgroundsInGhostWindow()

        ' The ghost window will be updated when it is resized

        Static Dim SizeChange As Integer = 1
        SizeChange = -SizeChange
        frmGhost.Size = New Size(frmGhostSkin.Size.Width + SizeChange, frmGhostSkin.Size.Height)

    End Sub

    Private Sub btnDefaults2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDefaults2.Click

        '0050 = Would you like to reset the this skin to its original defaults?

        Dim hold As Boolean = btnSave.Enabled

        If MsgBox(gResourceManager.GetString("0050"), MsgBoxStyle.YesNo Or MsgBoxStyle.Question, gResourceManager.GetString("0043")) = MsgBoxResult.Yes Then
            SetDefaultSkin()
            HoldSkin = gCurrentSkin
            btnSave.Enabled = hold
        End If

    End Sub

    Private Sub SetDefaultSkin()

        Loading = True ' this will surpress screen redraws

        SetCurrentSkinToGeneralDefaults()

        Me.cbEnabled.Checked = True
        Me.cbStretch.Checked = False

        Me.tbRotateHorizontal.Value = 1
        Me.tbRotateVertical.Value = 1
        Me.tbTransparency.Value = 100

        Loading = False ' resume screen redraws

        UpdateRulerBackgroundsInGhostWindow()
        Me.Refresh()

    End Sub

    Private Sub tbRotateHorizontal_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRotateHorizontal.ValueChanged

        If Loading Then Exit Sub

        btnSave.Enabled = True
        SetRotatedHorizontalBitmap()

    End Sub

    Private Sub tbRotateVertical_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRotateVertical.ValueChanged

        If Loading Then Exit Sub

        btnSave.Enabled = True
        SetRotatedVeritcalBitmap()

    End Sub

    Private Sub SetRotatedHorizontalBitmap()

        gCurrentSkin.HorizontalRotation = Me.tbRotateHorizontal.Value
        gCurrentSkin.HorizontalBackground = New Bitmap(gCurrentBaseBitMap)

        Select Case tbRotateHorizontal.Value

            Case Is = 1
                ' no rotation required

            Case Is = 2
                gCurrentSkin.HorizontalBackground.RotateFlip(RotateFlipType.Rotate90FlipNone)

            Case Is = 3
                gCurrentSkin.HorizontalBackground.RotateFlip(RotateFlipType.Rotate180FlipNone)

            Case Is = 4
                gCurrentSkin.HorizontalBackground.RotateFlip(RotateFlipType.Rotate270FlipNone)

        End Select

        UpdateRulerBackgroundsInGhostWindow()

    End Sub

    Private Sub SetRotatedVeritcalBitmap()

        gCurrentSkin.VerticalRotation = Me.tbRotateVertical.Value
        gCurrentSkin.VerticalBackground = New Bitmap(gCurrentBaseBitMap)

        Select Case tbRotateVertical.Value

            Case Is = 1
                ' no rotation required

            Case Is = 2
                gCurrentSkin.VerticalBackground.RotateFlip(RotateFlipType.Rotate90FlipNone)

            Case Is = 3
                gCurrentSkin.VerticalBackground.RotateFlip(RotateFlipType.Rotate180FlipNone)

            Case Is = 4
                gCurrentSkin.VerticalBackground.RotateFlip(RotateFlipType.Rotate270FlipNone)

        End Select

        UpdateRulerBackgroundsInGhostWindow()

    End Sub

    Private Sub LoadFirstSkin()

        Loading = True ' surpress screen redraws

        Dim WorkingSkinName As String = String.Empty

        For x As Integer = 0 To cbSkinNames.Items.Count - 1
            WorkingSkinName = TranslateASkinNameFromAnotherLanguageToEnglish(FilterSkinName(cbSkinNames.Items(x)))
            If (WorkingSkinName = gCurrentSkin.Name) Then
                cbSkinNames.SelectedIndex = x
                Exit For
            End If
        Next

        btnSave.Enabled = False
        MakeWindowMatchCurrentSkin()

        Loading = False ' resume screen redraws

        Me.Refresh()

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        SaveCurrentSkin(gCurrentSkin.Name)

        UpdateSkinNameDropDownBox()

        SetSkinsInStone()

        btnRemove.Enabled = (cbSkinNames.Items.Count > 1)

        MakeWindowMatchCurrentSkin()

        UpdateRulerBackgroundsInGhostWindow()

    End Sub

    Private Sub UpdateSkinNameDropDownBox()

        Loading = True

        Dim NewName As String = String.Empty
        Dim OldName As String = String.Empty

        If gCurrentSkin.SkinEnabled Then
            NewName = gCurrentSkin.Name & " " & gResourceManager.GetString("0044") ' (enabled)
            OldName = gCurrentSkin.Name & " " & gResourceManager.GetString("0045") ' (disabled)
        Else
            NewName = gCurrentSkin.Name & " " & gResourceManager.GetString("0045") ' (disabled)
            OldName = gCurrentSkin.Name & " " & gResourceManager.GetString("0044") ' (enabled)
        End If

        'Ditch the old name
        For x As Integer = 0 To cbSkinNames.Items.Count - 1
            If cbSkinNames.Items(x) = OldName Then
                cbSkinNames.Items.RemoveAt(x)
                Exit For
            End If
        Next

        'Add the new name (if it does't already exist)
        Dim NewNameAlreadyExists As Boolean = False
        For x As Integer = 0 To cbSkinNames.Items.Count - 1
            If cbSkinNames.Items(x) = NewName Then
                NewNameAlreadyExists = True
                Exit For
            End If
        Next

        If NewNameAlreadyExists Then
        Else
            cbSkinNames.Items.Add(NewName)
            cbSkinNames.Text = NewName
        End If

        Loading = False

    End Sub

    Private Sub cbSkinNames_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbSkinNames.SelectedValueChanged

        '0042 = "Remove (delete):"

        Me.ToolTip1.SetToolTip(Me.btnRemove, gResourceManager.GetString("0042") & " " & FilterSkinName(cbSkinNames.Items(cbSkinNames.SelectedIndex)))

        If Loading Then Exit Sub

        Dim HoldSelectedIndex As Integer = cbSkinNames.SelectedIndex

        Call ConfirmIfCurrentSkinShouldBeSaved()

        cbSkinNames.SelectedIndex = HoldSelectedIndex

        gCurrentSkin.Name = TranslateASkinNameFromAnotherLanguageToEnglish(FilterSkinName(cbSkinNames.Items(cbSkinNames.SelectedIndex)))
        cbSkinNames.Text = gCurrentSkin.Name

        btnSave.Enabled = False
        Call LoadSkin(gCurrentSkin.Name)

        HoldSkin = gCurrentSkin
        MakeWindowMatchCurrentSkin()

        MyTransparentColour = gCurrentSkin.TransparencyColour
        UpdateRulerBackgroundsInGhostWindow()

        btnSave.Enabled = False

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        If ConfirmIfCurrentSkinShouldBeSaved() Then
        Else
            gCurrentSkin = HoldSkin
            SetSkinsInStone()
        End If
        OkButtonPressed = True
        Me.Close()

    End Sub

    Private Function ConfirmIfCurrentSkinShouldBeSaved() As Boolean

        Dim ReturnCode As Boolean = False

        If btnSave.Enabled Then

            '0049 = Would you like to save your changes?
            If MsgBox(gResourceManager.GetString("0049"), MsgBoxStyle.YesNo Or MsgBoxStyle.Question, gResourceManager.GetString("0043")) = MsgBoxResult.Yes Then
                SaveCurrentSkin(gCurrentSkin.Name)
                UpdateSkinNameDropDownBox()
                SetSkinsInStone()
                ReturnCode = True
            End If

        End If

        Return ReturnCode

    End Function

    Private Sub SetSkinsInStone()

        RefreshAllSkinNames()
        SetNextAndPreviousEnabledSkinNames()
        HoldSkin = gCurrentSkin
        btnSave.Enabled = False

    End Sub

    Private Sub btnDefaultSkins_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDefaultSkins.Click

        Call ConfirmIfCurrentSkinShouldBeSaved()

        '0048 = Would you like to restore the original four skins (plastic see thru, stainless steel, wood, and yellow) skins to their original settings?
        If MsgBox(gResourceManager.GetString("0048"), MsgBoxStyle.YesNo Or MsgBoxStyle.Question, gResourceManager.GetString("0043")) = MsgBoxResult.Yes Then

            ResetOriginalSkins()
            LoadFirstSkin()
            SetSkinsInStone()
            btnRemove.Enabled = True

        End If

    End Sub

    Private Sub ResetOriginalSkins()

        ResetSkin("Wood")
        ResetSkin("Plastic see thru")
        ResetSkin("Yellow")
        ResetSkin("Stainless steel")

    End Sub

    Private Sub ResetSkin(ByVal SkinName As String)

        RemoveSkin(SkinName)

        gCurrentSkin.Name = SkinName
        cbSkinNames.Items.Add(TranslateASkinNameFromEnglishtToAnotherLanguage(SkinName) & " " & gResourceManager.GetString("0044")) '(enabled)
        SetCurrentSkinToGeneralDefaults()
        btnSave.Enabled = False
        SetDefaultBackgrounds(SkinName)
        SaveCurrentSkin(SkinName)

    End Sub

    Private Sub RemoveSkin(ByVal SkinName As String)

        SkinName = TranslateASkinNameFromAnotherLanguageToEnglish(SkinName)

        On Error Resume Next

        Dim XML_File_Name As String = XML_Path_Name & "\RulerDefinition_" & SkinName & MyExtention

        If System.IO.File.Exists(XML_File_Name) Then System.IO.File.Delete(XML_File_Name)

        For x As Integer = 0 To cbSkinNames.Items.Count - 1
            If (FilterSkinName(cbSkinNames.Items(x)) = SkinName) OrElse (FilterSkinName(cbSkinNames.Items(x)) = TranslateASkinNameFromEnglishtToAnotherLanguage(SkinName)) Then
                cbSkinNames.Items.RemoveAt(x)
                Exit For
            End If
        Next

    End Sub

    Private Function FilterSkinName(ByVal SkinName) As String

        Dim ReturnCode As String = String.Empty

        Try
            'ReturnCode = SkinName.Replace(" (enabled)", "").Replace(" (disabled)", "")
            ReturnCode = SkinName.Replace(" " & gResourceManager.GetString("0044"), "").Replace(" " & gResourceManager.GetString("0045"), "")

        Catch ex As Exception

        End Try
        Return ReturnCode

    End Function

    Private Sub btnLineBoxes_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles btnLineBoxes.Paint, btnLineLength.Paint, btnLineMeasuring.Paint, btnLineMidpoint.Paint, btnLineTicksAndSides.Paint, btnNumbers.Paint, btnLineGoldenRatio.Paint, btnLineThirds.Paint

        Dim WorkingColour As Color

        With gCurrentSkin

            Select Case sender.name

                Case Is = "btnLineBoxes"
                    WorkingColour = .LineBoxes

                Case Is = "btnLineTicksAndSides"
                    WorkingColour = .LineTicksAndSides

                Case Is = "btnLineLength"
                    WorkingColour = .LineLength

                Case Is = "btnLineMeasuring"
                    WorkingColour = .LineMeasuring

                Case Is = "btnLineMidpoint"
                    WorkingColour = .LineMidpoint

                Case Is = "btnLineThirds"
                    WorkingColour = .LineThirds

                Case Is = "btnLineGoldenRatio"
                    WorkingColour = .LineGoldenRatio

                Case Is = "btnLineTicks"
                    WorkingColour = .LineTicksAndSides

                Case Is = "btnNumbers"
                    WorkingColour = .Numbers

            End Select

        End With

        e.Graphics.DrawLine(New Pen(WorkingColour), 10 * gUniversaleScale, 17 * gUniversaleScale, 98 * gUniversaleScale, 17 * gUniversaleScale)

    End Sub

    Private Sub cbEnabled_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbEnabled.CheckedChanged

        '0046 = There must be at least one enabled skin.  As this is the only enabled skin it cannot be disabled.
        '0047 = A Ruler for Windows - Info

        If Loading Then Exit Sub

        If cbEnabled.Checked Then
        Else
            If ThereAreAtLeastTwoEnabledSkins() Then
            Else
                MsgBox(gResourceManager.GetString("0046"), MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, gResourceManager.GetString("0047")) : Me.DialogResult = 0
                Loading = True
                cbEnabled.Checked = True
                Loading = False
                Exit Sub
            End If
        End If

        gCurrentSkin.SkinEnabled = cbEnabled.Checked
        btnSave.Enabled = True

    End Sub

    Private Sub cbStretch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStretch.CheckedChanged

        If Loading Then Exit Sub

        gCurrentSkin.Tiled = cbStretch.Checked
        UpdateRulerBackgroundsInGhostWindow()
        btnSave.Enabled = True

    End Sub

    Private Sub btnSave_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.EnabledChanged

        If sender.enabled Then
            btnSave.ForeColor = Color.White
            btnSave.BackColor = Color.Green
        Else
            btnSave.ForeColor = Color.Black
            btnSave.BackColor = Color.White
        End If

    End Sub

    Private Function ThereAreAtLeastTwoEnabledSkins() As Boolean

        Dim CountofEnabledSkins As Integer = 0
        RefreshAllSkinNames()

        For Each SkinName As String In gAllSkinNames
            If GetEnabledSetting(SkinName) Then CountofEnabledSkins += 1
            If CountofEnabledSkins > 1 Then Return True
        Next

        Return False

    End Function

    Private Sub btnLocate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLocate.Click

        Dim procstart As New ProcessStartInfo("explorer")
        procstart.Arguments = XML_Path_Name
        Process.Start(procstart)

        'don't ask me why, but once the thread starts the ghost window needs to be redrawn
        Threading.Thread.Sleep(1500)
        Application.DoEvents()
        UpdateRulerBackgroundsInGhostWindow()
        Application.DoEvents()

    End Sub

    Private Sub btnGetMoreFreeSkinsOnline_Click(sender As Object, e As EventArgs) Handles btnGetMoreFreeSkinsOnline.Click

        StartAProcess(gResourceManager.GetString("0066")) 'skins website
        Me.Close()

    End Sub

    Private Sub frmSkins_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        gSuspendTimer2 = hgSuspendTimer2
    End Sub

End Class