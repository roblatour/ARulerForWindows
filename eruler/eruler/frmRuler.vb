Imports System.ComponentModel
Imports System.IO
Imports System.Configuration
Imports System.Linq
Public Class frmRuler

    Private oldMagnifierFactor As Integer = 1
    Private RulerHasScaleOnTopOrLeft As Boolean = True

    Private g As Graphics

    Private np As New Point

    Private NewRuler, AboutBox, ResizeLeftOrUpBox, ResizeRightOrDownBox, NumberBox, MinimizeNowBox, MagnifyBox, TickPositionBox, HorizontalVeriticalBox, CloseBox, BackgroundBox, MeasuringBox, MasterBoxA, MasterBoxB As Rectangle

    Private MoveFlag As Boolean = False
    Private ResizeRightFlag As Boolean = False
    Private ResizeLeftFlag As Boolean = False

    Private AbsoluteMousePositionOnRulerX As Integer = 0
    Private AbsoluteMousePositionOnRulerY As Integer = 0

    Private RelativeMousePositionOnRulerX As Integer
    Private RelativeMousePositionOnRulerY As Integer

    'Measuring Line
    Private MeasuringLineX As Integer = -1
    Private MeasuringLineY As Integer = -1
    Private MeasuringLineXComplement As Integer = -1
    Private MeasuringLineYComplement As Integer = -1

    'Midpoint
    Private MidpointLineX As Integer = -1
    Private MidpointLineY As Integer = -1
    Private MidpointLineXComplement As Integer = -1
    Private MidpointLineYComplement As Integer = -1

    Private WidthInPixels As Single 'Integer
    Private HeightInPixels As Single 'Integer
    Private AdjustReadingGuideThickness As Integer

    Private ClickToRight As Integer
    Private ClickToLeft As Integer

    Private AutoScrollPositionX As Integer
    Private AutoScrollPositionY As Integer

    Private NumberingIsLeftToRightHorizontal As Boolean = True
    Private NumberingIsSmallOnTopToHighOnBottomVertical As Boolean = True

    Private MidPointIsOn As Boolean = False
    Private ThirdsAreOn As Boolean = False
    Private GoldenRatioIsOn As Boolean = False

    Private HoldMidPointIsOn As Boolean = False
    Private HoldRedLineX As Integer = -1
    Private HoldRedLineXComplement As Integer = -1
    Private HoldRedLineY As Integer = -1
    Private HoldRedLineYComplement As Integer = -1

    Private Enum CommandLineResults
        NoSkin = 0
        GoodSkin = 1
        BadSkin = 2
        ExistingSkin = 3
    End Enum

    Private FirstActivation As Boolean = True

    'Const gBigBox As Integer = 20
    'Const gNormalBox As Integer = 10
    'Private gBoxScale = gNormalBox
    Private gOldBoxScale = gNormalBox
    Private gCollapseCount As Integer = 0
    Private gHoldShowLength As Boolean
    Private gHoldShowLengthSet As Boolean = False

    Private gOpenedBasedOnCommandLineRequest As Boolean = False

    Private Sub FrmRuler_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            SetSettingsVersion()

            SetLanguage()

            ' command line overrides

            Dim CommandLine As String = Microsoft.VisualBasic.Command.Trim.ToLower
            ' CommandLine = "mode=ruler or=horizontal width=717 location=(1390,444)"

            'Inventory active screens
            gScreens = Screen.AllScreens
            For x As Int16 = 0 To gScreens.Length - 1
                gMinX = Math.Min(gMinX, gScreens(x).WorkingArea.X)
                gMinY = Math.Min(gMinY, gScreens(x).WorkingArea.Y)
                gMaxWidth += gScreens(x).Bounds.Width
                gMaxHeight += gScreens(x).Bounds.Height
            Next

            gPrimaryScreenHome = New Point(-gMinX, -gMinY)

            ' find out if the program is already running
            Dim aModuleName As String = Diagnostics.Process.GetCurrentProcess.MainModule.ModuleName
            Dim aProcName As String = System.IO.Path.GetFileNameWithoutExtension(aModuleName)
            Dim NumberOfOtherARulersForWindowsThatAreRunning As Integer = Process.GetProcessesByName(aProcName).Length - 1

            Dim targetx As Integer = My.Settings.LocationX + NumberOfOtherARulersForWindowsThatAreRunning * 20
            Dim targety As Integer = My.Settings.LocationY + NumberOfOtherARulersForWindowsThatAreRunning * 20


            Dim Scale As Single = CSng(1)

            If CommandLine.Contains("scale=true") Then

                ' apply scale to be applied to location, width and height
                gOpenedBasedOnCommandLineRequest = True

                Dim DPI As Integer = SafeNativeMethods.GetDpiForWindow(Me.Handle)
                Scale = CSng(DPI) / CSng(96)

            End If

            Try

                If CommandLine.Contains("location=(") Then

                    gOpenedBasedOnCommandLineRequest = True

                    Dim ws As String = CommandLine
                    ws = ws.Remove(0, ws.IndexOf("location=(") + "location=(".Length)

                    targetx = CInt(ws.Remove(ws.IndexOf(",")))
                    ws = ws.Remove(0, ws.IndexOf(",") + 1)
                    targety = CInt(ws.Remove(ws.IndexOf(")")))

                    If Scale = CSng(1) Then

                        targetx -= gMinX
                        targety -= gMinY

                    Else

                        targetx = CInt(CSng(targetx) * Scale) - gMinX
                        targety = CInt(CSng(targety) * Scale) - gMinY

                        'MsgBox("scale " & Scale & vbCrLf &
                        '       "gMinX " & gMinX & vbCrLf &
                        '       "gMinY " & gMinY & vbCrLf &
                        '       "targetx " & targetx & vbCrLf &
                        '       "targety " & targety & vbCrLf &
                        '       "ntargetx " & ntargetx & vbCrLf &
                        '       "ntargety " & ntargety)

                    End If

                End If

            Catch ex As Exception

            End Try

            If targetx > gMaxWidth - WidthInPixels Then
                targetx = NumberOfOtherARulersForWindowsThatAreRunning * 20
            End If

            If targety > gMaxHeight - HeightInPixels Then
                targety = NumberOfOtherARulersForWindowsThatAreRunning * 20
            End If


            RulerPanel.Location = New Point(targetx, targety)


            If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
            Else
                My.Settings.LocationX = gPrimaryScreenHome.X
                My.Settings.LocationY = gPrimaryScreenHome.Y
                My.Settings.Horizontal = True
                My.Settings.Width = 250
                My.Settings.Save()
                RulerPanel.Location = New Point(gPrimaryScreenHome.X, gPrimaryScreenHome.Y)
            End If

            If My.Settings.scale = 0.0625 Then
            Else
                If (My.Settings.scale > 5) OrElse (My.Settings.scale < 0.5) OrElse (My.Settings.scale = 0) Then
                    My.Settings.scale = 1
                    My.Settings.Save()
                End If
            End If


            Me.AutoScroll = False

            Me.BackColor = MyTransparentColour
            Me.TransparencyKey = MyTransparentColour

            'Load Settings
            gShowBalloons = My.Settings.ShowHelpBalloonsOnRuler
            gQuickStartup = My.Settings.QuickStartup
            gFenceRulerOnScreen = My.Settings.FenceRulerOnScreen
            gAutoExpand = My.Settings.AutoExpand
            gBump = My.Settings.Bump
            gShowLength = My.Settings.ShowLength
            gShowReadingGuide = My.Settings.ShowReadingGuide
            gRulerScalingFactorSetByUser = My.Settings.scale

            RulerIsHorizontal = My.Settings.Horizontal
            RulerIsVertical = Not RulerIsHorizontal
            RulerHasScaleOnTopOrLeft = My.Settings.ScaleOnTop

            NumberingIsLeftToRightHorizontal = My.Settings.LeftToRightHor
            NumberingIsSmallOnTopToHighOnBottomVertical = My.Settings.LeftToRightVer

            AdjustReadingGuideThickness = My.Settings.AdjustReadingGuideThickness

            If CommandLine.Contains("mode=rg") Then
                gShowReadingGuide = True
                gOpenedBasedOnCommandLineRequest = True

            ElseIf CommandLine.Contains("mode=ruler") Then
                gShowReadingGuide = False
                gOpenedBasedOnCommandLineRequest = True

            End If

            If CommandLine.Contains("or=horizontal") Then
                RulerIsHorizontal = True
                gOpenedBasedOnCommandLineRequest = True

            ElseIf CommandLine.Contains("or=vertical") Then
                RulerIsHorizontal = False
                gOpenedBasedOnCommandLineRequest = True

            End If


            Dim lWidth As Integer = My.Settings.Width

            If CommandLine.Contains("width=") Then

                gOpenedBasedOnCommandLineRequest = True

                Dim ws As String = CommandLine
                ws = ws.Remove(0, ws.IndexOf("width=") + "width=".Length)

                Dim j As Integer = ws.IndexOf(" ")

                If ws.IndexOf(" ") > -1 Then
                    ws = ws.Remove(ws.IndexOf(" ")).Trim
                End If

                Try
                    lWidth = CInt(ws) * Scale
                Catch ex As Exception
                    lWidth = My.Settings.Width
                End Try

            End If

            Dim lHeight As Integer = My.Settings.Height

            If CommandLine.Contains("height=") Then

                gOpenedBasedOnCommandLineRequest = True

                Dim ws As String = CommandLine
                ws = ws.Remove(0, ws.IndexOf("height=") + "height=".Length)

                If ws.IndexOf(" ") > -1 Then
                    ws = ws.Remove(ws.IndexOf(" ")).Trim
                End If

                Try
                    lHeight = CInt(ws) * Scale
                Catch ex As Exception
                    lHeight = My.Settings.Height
                End Try

            End If


            If CommandLine.Contains("rglines=false") Then
                gNoLines = True
            End If
#If DEBUG Then
            '  gNoLines = True
#End If


            XML_Path_Name = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) & "\ARulerForWindows"

            EstablishCurrentSkin()

            Me.Location = New Point(gMinX, gMinY)

            Me.Width = My.Computer.Screen.WorkingArea.Width
            Me.Height = My.Computer.Screen.WorkingArea.Height

            RulerPanel.Visible = False
            MagnifyViewPanel.Visible = False

            If RulerIsHorizontal Then
                If lWidth = 0 Then
                    WidthInPixels = My.Computer.Screen.WorkingArea.Width / 2
                Else
                    WidthInPixels = Math.Min(lWidth, My.Computer.Screen.WorkingArea.Width)
                End If
                HeightInPixels = 100 'key 
            Else
                If lHeight = 0 Then
                    HeightInPixels = My.Computer.Screen.WorkingArea.Height / 2
                Else
                    HeightInPixels = Math.Min(lHeight, My.Computer.Screen.WorkingArea.Height)
                End If
                WidthInPixels = 100 'key
            End If

            g = RulerPanel.CreateGraphics

            gUniversaleScale = g.DpiX / 96 'testing here
            NumbersFont = New Font(NumbersFont.FontFamily, NumbersFont.Size / gUniversaleScale)
            SmallNumbersFont = New Font(SmallNumbersFont.FontFamily, SmallNumbersFont.Size / gUniversaleScale)

            MakeTopMost(Me.Handle.ToInt64)

            SetupForRulerOrReadingGuide()

            Me.RulerPanel.Visible = True

            RulerDrawBackground()

            MagnifyViewPanel.Location = New Point(0, 0)
            MagnifyViewPanel.Width = Screen.PrimaryScreen.WorkingArea.Width
            MagnifyViewPanel.Height = Screen.PrimaryScreen.WorkingArea.Height

            Timer1.Enabled = True
            Timer1.Start()

            If gQuickStartup Then
                ShortCut("Ensure")
            End If

            If Microsoft.VisualBasic.Command.ToString.ToUpper.Contains("QUICKSTART") Then
                Me.Size = New Size(1, 1)
                Me.Location = New Point(-500, -500)
                Application.Exit()
            Else
                CreateAMutex()
                Flush()
            End If


            'if there is another program running, don't agressively try and keep this one on top.
            If NumberOfOtherARulersForWindowsThatAreRunning = 0 Then
                Timer2.Interval = 500
                Timer2.Enabled = True
                Timer2.Start()
            End If

            If gOpenedBasedOnCommandLineRequest Then
                SetupForFileWatcherMonitoring()
            End If

            ' v3.5
            ' added in v3.5 allow ruler to be fully transparen for user specified rgb value
            ' RulerPanel.BackColor = Color.FromArgb(My.Settings.TransparencyA, My.Settings.TransparencyR, My.Settings.TransparencyG, My.Settings.TransparencyB)

            ' RulerPanel.BackColor = Color.Transparent
            ' Console.WriteLine(Color.FromArgb(255, 250, 250, 250).ToArgb.ToString)

            SetupForMagnifier()

        Catch ex As Exception
            MsgBox(ex.TargetSite.Name & vbCrLf & ex.Message, MsgBoxStyle.Critical, gResourceManager.GetString("0067")) : Me.DialogResult = 0
        End Try

    End Sub


    Private gMagnifierFound As Boolean

    Private gMagnifierProgram As String = String.Empty

    Private W10_32_MagnifierProgram As String = "C:\Windows\System32\Magnify.exe"
    Private W10_64_MagnifierProgram As String = "c:\Windows\Sysnative\Magnfiy.exe"
    Private Sub SetupForMagnifier()

        If File.Exists(W10_64_MagnifierProgram) Then
            gMagnifierProgram = W10_64_MagnifierProgram
        ElseIf File.Exists(W10_32_MagnifierProgram) Then
            gMagnifierProgram = W10_32_MagnifierProgram
        End If

        gMagnifierFound = (gMagnifierProgram.Length > 0)

    End Sub


    Private Const gFileWatcherFilterName As String = "ar4w_close.txt"
    Private gFileWatcherPathName As String = String.Empty
    Private gFileWatcherPathAndFileName As String = String.Empty
    Private gCloseWithoutPrompt As Boolean = False
    Private Sub SetupForFileWatcherMonitoring()

        Try

            gFileWatcherPathName = System.IO.Path.GetTempPath()

            gFileWatcherPathAndFileName = gFileWatcherPathName & gFileWatcherFilterName

            Dim fw As New FileSystemWatcher
            fw.Path = gFileWatcherPathName
            fw.IncludeSubdirectories = False
            fw.Filter = gFileWatcherFilterName
            AddHandler fw.Created, New FileSystemEventHandler(AddressOf FileWatcherDetectedAFileWasCreated)
            fw.EnableRaisingEvents = True

        Catch ex As Exception

        End Try

    End Sub


    Friend Sub FileWatcherDetectedAFileWasCreated(ByVal source As Object, ByVal e As FileSystemEventArgs)

        gCloseWithoutPrompt = True
        FrmRuler_Close()

    End Sub

    Private Sub LoadANewSkinFromTheCommandLine()

        Dim ErrorFound As Boolean = True

        Try

            'Get parameters passed into this program

            Dim InputFileName As String = Microsoft.VisualBasic.Command.ToString.Trim
            If InputFileName.Length = 0 Then Exit Sub ' no parms were passed in

            If InputFileName.ToUpper.Contains("QUICKSTART") Then Exit Sub ' quick startup logic

            If InputFileName.ToLower.Contains(".ar4w") Then
            Else
                Exit Sub
            End If
            '***************************************************************************

            Dim CandidateSkin As String = InputFileName

            'Ensure the file has the right file name format
            If CandidateSkin.Length <= (Len("\RulerDefinition_.ar4w") + 2) Then Exit Try

            If CandidateSkin.Contains(".") Then
            Else
                Exit Try
            End If

            If CandidateSkin.StartsWith("""") Then
                CandidateSkin = CandidateSkin.Remove(0, 1)
            Else
                Exit Try
            End If

            If CandidateSkin.EndsWith("""") Then
                CandidateSkin = CandidateSkin.Remove(CandidateSkin.Length - 1, 1)
            Else
                Exit Try
            End If

            InputFileName = CandidateSkin ' same as before but with the quotes removed

            Dim ext As String = System.IO.Path.GetExtension(CandidateSkin)
            If (String.Compare(ext, MyExtention, True) <> 0) Then Exit Try

            If CandidateSkin.ToUpper.EndsWith(MyExtention.ToUpper) Then
                CandidateSkin = CandidateSkin.Remove(CandidateSkin.Length - 5)
            Else
                Exit Try
            End If

            If CandidateSkin.ToUpper.Contains("\RULERDEFINITION_") Then
            Else
                Exit Try
            End If

            'Me.Location = gPrimaryScreenHome ' testing here

            CandidateSkin = CandidateSkin.Remove(0, Command.LastIndexOf("\RulerDefinition_") + Len("\RulerDefinition_") - 1)

            'Convert all %20s to spaces
            CandidateSkin = CandidateSkin.Replace("%20", " ")

            ' Remove the [1], [2], etc. that may have been tacked onto the end of the candiate skin name by helpful browser software
            Dim TestForIEAddedExtention As String

            For x As Integer = 1 To 25

                TestForIEAddedExtention = "[" & CType(x, String).Trim & "]"
                If CandidateSkin.EndsWith(TestForIEAddedExtention) Then
                    CandidateSkin = CandidateSkin.Remove(CandidateSkin.LastIndexOf(TestForIEAddedExtention))
                End If

                TestForIEAddedExtention = "(" & CType(x, String).Trim & ")"
                If CandidateSkin.EndsWith(TestForIEAddedExtention) Then
                    CandidateSkin = CandidateSkin.Remove(CandidateSkin.LastIndexOf(TestForIEAddedExtention))
                End If
            Next

            CandidateSkin = CandidateSkin.Trim

            Dim XML_File_Name As String = XML_Path_Name & "\RulerDefinition_" & CandidateSkin & MyExtention

            'Check to see if that name is already in use

            ErrorFound = False
            Dim LoadRequired As Boolean = False

            If System.IO.File.Exists(XML_File_Name) Then

                ' "That skin name already exists." "Would you Like to replace it with this one?
                If MsgBox(gResourceManager.GetString("0074") & vbCrLf &
                          gResourceManager.GetString("0075"),
                           MsgBoxStyle.YesNo Or MsgBoxStyle.Question, gResourceManager.GetString("0067")) = MsgBoxResult.Yes _
                           Then
                    DeleteSkin(CandidateSkin)
                    LoadRequired = True
                End If

            Else
                LoadRequired = True
            End If


            If LoadRequired Then

                System.IO.File.Copy(InputFileName, XML_File_Name)
                ConfirmCurrentSkinHasReadWriteAuthority(XML_File_Name)

                If LoadSkin(CandidateSkin) Then

                    gCurrentSkin.Name = CandidateSkin

                    My.Settings.CurrentSkinName = CandidateSkin
                    My.Settings.Save()

                    RefreshAllSkinNames()

                    SetNextAndPreviousEnabledSkinNames()

                    Me.BackColor = gCurrentSkin.TransparencyColour
                    Me.TransparencyKey = gCurrentSkin.TransparencyColour
                    Me.Opacity = gCurrentSkin.Opacity

                    RulerDrawBackground()

                Else

                    ErrorFound = True

                End If

            End If


        Catch ex As Exception

            ' 0067 = A Ruler for Windows - Error
            ' 0068 = could not load the skin

            'MsgBox(ex.ToString)
            MsgBox(gResourceManager.GetString("0068"), MsgBoxStyle.Critical, gResourceManager.GetString("0067")) : Me.DialogResult = 0

        End Try

        If ErrorFound Then
            ' 0067 = A Ruler for Windows - Error
            '0069 = Problem trying to load the .ar4w file
            MsgBox(gResourceManager.GetString("0068"), MsgBoxStyle.Critical, gResourceManager.GetString("0067")) : Me.DialogResult = 0
        End If

    End Sub

    Private Sub frmRuler_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If FirstActivation Then
            FirstActivation = False
            LoadANewSkinFromTheCommandLine()
        End If

    End Sub

    Private Sub EstablishCurrentSkin()

        Try

            CreateFullDirectory(XML_Path_Name)

            gCurrentSkin.Name = String.Empty

            RefreshAllSkinNames()

            If gAllSkinNames Is Nothing Then

                ResetSkin("Wood")
                ResetSkin("Plastic see thru")
                ResetSkin("Yellow")
                ResetSkin("Stainless steel")

                gCurrentSkin.Name = "Wood"
                gNextEnabledSkinName = "Yellow"

                RefreshAllSkinNames()

            End If

            For x As Integer = 0 To gAllSkinNames.Length - 1

                If gAllSkinNames(x) = My.Settings.CurrentSkinName Then

                    gCurrentSkin.Name = gAllSkinNames(x)
                    SetNextAndPreviousEnabledSkinNames()

                End If

            Next

            If gCurrentSkin.Name.Length = 0 Then

                'catches if the current skin name is for some reason not available

                gCurrentSkin.Name = gAllSkinNames(0)

                If gAllSkinNames.Length = 1 Then
                    gNextEnabledSkinName = gCurrentSkin.Name
                Else
                    gNextEnabledSkinName = gAllSkinNames(1)
                End If

            End If

            Call LoadSkin(gCurrentSkin.Name)
            SetNextAndPreviousEnabledSkinNames()

            Me.BackColor = gCurrentSkin.TransparencyColour
            Me.TransparencyKey = gCurrentSkin.TransparencyColour
            Me.Opacity = gCurrentSkin.Opacity

        Catch ex As Exception

            MsgBox(gResourceManager.GetString("0062") & ex.ToString) ' "Problem establishing current skin"

        End Try

    End Sub

    Private Sub ResetSkin(ByVal SkinName As String)

        gCurrentSkin.Name = String.Empty
        SetCurrentSkinToGeneralDefaults()

        gCurrentSkin.Name = SkinName
        gCurrentSkin.SkinEnabled = True
        SetDefaultBackgrounds(gCurrentSkin.Name)
        SaveCurrentSkin(gCurrentSkin.Name)

    End Sub

    Private Sub FrmRuler_Close()

        If CloseOnlyOnce OrElse CloseOnDemand Then
            CloseOnlyOnce = False
            CloseOnDemand = False
        Else
            Exit Sub
        End If

        If gCloseWithoutPrompt Then
        Else

            If My.Settings.ConfirmOnExit Then

                If MagnifierFactor = 1 Then
                    gSetLocation = New Point(RulerPanel.Location.X + gMinX, RulerPanel.Location.Y + gMinY)
                Else
                    gSetLocation = New Point(RulerPanel.Location.X + gMinX + gPrimaryScreenHome.X, RulerPanel.Location.Y + gMinY + gPrimaryScreenHome.Y)
                End If

                Dim frmConfirmExit As Form = New frmConfirmExit
                frmConfirmExit.ShowDialog()
                frmConfirmExit.Dispose()

                If Not gOKToExitNow Then
                    RulerDrawBackground()
                    Exit Sub
                End If

            End If

        End If

        Me.Hide()
        Application.DoEvents()

        My.Settings.Horizontal = RulerIsHorizontal

        My.Settings.LeftToRightHor = NumberingIsLeftToRightHorizontal
        My.Settings.LeftToRightVer = NumberingIsSmallOnTopToHighOnBottomVertical

        My.Settings.ScaleOnTop = RulerHasScaleOnTopOrLeft
        My.Settings.LocationX = RulerPanel.Location.X
        My.Settings.LocationY = RulerPanel.Location.Y

        My.Settings.Width = RulerPanel.Width
        My.Settings.Height = RulerPanel.Height
        My.Settings.ShowHelpBalloonsOnRuler = gShowBalloons
        My.Settings.FenceRulerOnScreen = gFenceRulerOnScreen
        My.Settings.AutoExpand = gAutoExpand
        My.Settings.Bump = gBump
        My.Settings.ShowLength = gShowLength
        My.Settings.ShowReadingGuide = gShowReadingGuide
        My.Settings.AdjustReadingGuideThickness = AdjustReadingGuideThickness

        My.Settings.CurrentSkinName = gCurrentSkin.Name

        My.Settings.Save()
        Application.DoEvents()

        MakeTopMost(Me.Handle.ToInt64, False)
        Application.DoEvents()

        System.Threading.Thread.Sleep(1500)

        Me.Close()
        Application.DoEvents()
        End

    End Sub

    Const VK_LControl As Integer = &HA2
    Const VK_Shift As Integer = &H10

    'Symbolic         Hex        Dec      Keyboard
    'Constant         Val        Val      Equivalent
    'VK_SHIFT         0x10        16         SHIFT
    'VK_LSHIFT        0xA0       160      L. SHIFT
    'VK_RSHIFT        0xA1       161      R. SHIFT

    'VK_CONTROL       0x11        17         CTRL
    'VK_LCONTROL      0xA2       162      L. CTRL 
    'VK_RCONTROL      0xA3       163      R. CTRL

    'VK_MENU          0x12        18         ALT
    'VK_LMENU         0xA4       164      L. ALT
    'VK_RMENU         0xA5       165      R. ALT

    'VK_PAUSE         0x13        19         PAUSE
    'VK_ESCAPE        0x1B        27         ESC
    'VK_RETURN        0x0D         13        ENTER
    'VK_CAPITAL       0x14        20         CAPS LOCK
    'VK_NUMLOCK       0x90       144         NUM LOCK
    'VK_SCROLL        0x91       145         SCROLL LOCK


    'adding in version 4.2.7
    Private Sub Form1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel

        Dim e1 As System.Windows.Forms.KeyEventArgs = New System.Windows.Forms.KeyEventArgs(Keys.B) 'e1 set to an unused key

        Dim ClearTheBuffer As Integer = SafeNativeMethods.GetAsyncKeyState(VK_Shift)

        If SafeNativeMethods.GetAsyncKeyState(VK_Shift) Then

            If e.Delta > 0 Then
                ScrollRight(e1)
            Else
                ScrollLeft(e1)
            End If

        Else

            If e.Delta > 0 Then
                ScrollUp(e1)
            Else
                ScrollDown(e1)
            End If

        End If

    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Dim ResetSizeBoxes As Boolean = False

        Try

            MoveFlag = False

            Dim AnActionKeyWasPressed As Boolean = True

            Select Case e.KeyCode

                Case Keys.Left

                    ScrollLeft(e)

                Case Keys.Right

                    ScrollRight(e)

                Case Keys.Up

                    ScrollUp(e)

                Case Keys.Down, 32 ' spacebar

                    ScrollDown(e)

                Case Keys.ControlKey

                    ' Left and Right Control Keys lock the ruler
                    If SafeNativeMethods.GetAsyncKeyState(VK_LControl) Then
                        If RulerIsHorizontal Then
                            gLockLeft = Not gLockLeft
                        Else
                            gLockTop = Not gLockTop
                        End If
                    Else
                        If RulerIsHorizontal Then
                            gLockRight = Not gLockRight
                        Else
                            gLockBottom = Not gLockBottom
                        End If
                    End If

                    SetBoxSizes()

                Case Keys.Home

                    If MagnifierFactor > 1 Then RestoreNormalScreen()

                    ClearRedMeasuringLine()

                    MidPointIsOn = False

                    gShowLength = True

                    RulerPanel.Location = gPrimaryScreenHome

                    gShowReadingGuide = False
                    SetupForRulerOrReadingGuide()

                    ControlsAreHidden = False

                    RulerIsHorizontal = True
                    RulerIsVertical = False
                    RulerHasScaleOnTopOrLeft = True

                    WidthInPixels = 250
                    HeightInPixels = 100

                    MeasuringLineX -= 1
                    MeasuringLineXComplement -= 1
                    MeasuringLineY -= 1
                    MeasuringLineYComplement -= 1

                    MidPointIsOn = False

                    NumberingIsLeftToRightHorizontal = True
                    NumberingIsSmallOnTopToHighOnBottomVertical = True

                    SetBoxSizes()

                Case Keys.Escape
                    ' Escape either exits from Magnify mode or exits the program

                    If MagnifierFactor > 1 Then
                        RestoreNormalScreen()
                        Exit Sub
                    Else
                        CloseOnDemand = True
                        FrmRuler_Close()
                        Exit Sub
                    End If

                Case 51 ' "#"
                    ShowResizeNow()

                Case Is = Keys.I
                    If e.Shift OrElse gShowReadingGuide Then
                        gShowReadingGuide = Not gShowReadingGuide
                        SetupForRulerOrReadingGuide()
                    Else
                        SwapTicksNow()
                    End If

                Case Is = Keys.O
                    If e.Shift Then
                        ChangeBackgroundNow(True)
                    Else
                        ChangeBackgroundNow(False)
                    End If

                Case Keys.X
                    CloseOnDemand = True
                    FrmRuler_Close()

                Case 107, 187 ' "+"
                    MagnifyNow()

                Case 109 ' "-"
                    MinimizeNow()

                Case 111 '"/"

                    FlipVerticalHorizontal()
                    If gShowReadingGuide Then
                        SetupForRulerOrReadingGuide()
                    End If

                Case Is = 189 ' "-"
                    If Not e.Shift Then MinimizeNow()

                Case Is = Keys.OemQuestion ' "?" 'v1.7 (Dutch (nl) keyboards)

                    If My.Settings.Language.StartsWith("nl") Then

                        If e.Shift Then
                            FlipVerticalHorizontal()
                            If gShowReadingGuide Then
                                SetupForRulerOrReadingGuide()
                            End If
                        End If

                    Else

                        If e.Shift Then
                            ShowHelpNow()
                        Else
                            FlipVerticalHorizontal()
                            If gShowReadingGuide Then
                                SetupForRulerOrReadingGuide()
                            End If
                        End If

                    End If

                Case Is = Keys.Oemcomma ' "?" 'v1.7 (Dutch (nl) keyboards)

                    If My.Settings.Language.StartsWith("nl") Then

                        If e.Shift Then
                            ShowHelpNow()
                        End If

                    End If

                Case Is = Keys.PageUp 'v3.3.3

                    If gShowReadingGuide Then

                        If RulerIsHorizontal Then
                            AnActionKeyWasPressed = MakeReadingGuideThinner()
                        Else
                            AnActionKeyWasPressed = MakeReadingGuideThicker()
                        End If

                    End If

                Case Is = Keys.PageDown 'v3.3.3

                    If gShowReadingGuide Then

                        If RulerIsHorizontal Then
                            AnActionKeyWasPressed = MakeReadingGuideThicker()
                        Else
                            AnActionKeyWasPressed = MakeReadingGuideThinner()
                        End If

                    End If


                Case Else

                    ' The following keys are language sensitive

                    Dim KeyPressed = UCase(e.KeyCode.ToString)

                    'Console.WriteLine(KeyPressed)

                    Select Case KeyPressed

                        Case Is = gKey_Clear

                            ClearRedMeasuringLine()

                        Case Is = gKey_ShowGoldenRatioLine

                            GoldenRatioIsOn = Not GoldenRatioIsOn

                        Case Is = gKey_ShowLength

                            gShowLength = Not gShowLength

                        Case Is = gKey_ShowMidPoint

                            MidPointIsOn = Not MidPointIsOn

                        Case Is = gKey_ShowThirds

                            ThirdsAreOn = Not ThirdsAreOn

                        Case Is = gKey_ReverseNumbers

                            If RulerIsHorizontal Then
                                swap(MeasuringLineX, MeasuringLineXComplement)
                                swap(MidpointLineX, MidpointLineXComplement)
                                NumberingIsLeftToRightHorizontal = Not NumberingIsLeftToRightHorizontal
                            Else
                                swap(MeasuringLineY, MeasuringLineYComplement)
                                swap(MidpointLineY, MidpointLineYComplement)
                                NumberingIsSmallOnTopToHighOnBottomVertical = Not NumberingIsSmallOnTopToHighOnBottomVertical
                            End If

                        Case Is = gKey_SwitchBetweenRulerAndGuide

                            gShowReadingGuide = Not gShowReadingGuide
                            SetupForRulerOrReadingGuide()

                            ControlsAreHidden = False
                            Timer1.Enabled = True
                            Timer1.Start()

                        Case Else

                            AnActionKeyWasPressed = False

                    End Select

            End Select

            If AnActionKeyWasPressed Then
                RulerDrawBackground()
            End If

        Catch ex As Exception
            MsgBox(ex.TargetSite.Name & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error") : Me.DialogResult = 0
        End Try

        MoveFlag = False

    End Sub

    Private Function MakeReadingGuideThicker() As Boolean

        Static LastAdjustment As DateTime = Now
        Static LastPass As Double = 1

        Dim ReturnCode As Boolean

        Dim OneHalfOfTheCurrentScreenHeight As Integer = FindHeightOfCurrentScreen(RulerPanel.Location) / 2

        If AdjustReadingGuideThickness < OneHalfOfTheCurrentScreenHeight Then

            Dim ts As TimeSpan = Now - LastAdjustment
            Dim diff As Integer = Int(ts.TotalMilliseconds)

            If diff < 100 Then
                LastPass += 0.5
            Else
                LastPass = 1
            End If

            AdjustReadingGuideThickness += LastPass

            LastAdjustment = Now

            If AdjustReadingGuideThickness > OneHalfOfTheCurrentScreenHeight Then AdjustReadingGuideThickness = OneHalfOfTheCurrentScreenHeight

            SetupForRulerOrReadingGuide()

            ReturnCode = True

        Else

            ReturnCode = False

        End If

        Return ReturnCode

    End Function

    Private Function MakeReadingGuideThinner() As Boolean

        Static LastAdjustment As DateTime = Now
        Static LastPass As Double = 1

        Dim ReturnCode As Boolean

        If AdjustReadingGuideThickness > 0 Then

            Dim ts As TimeSpan = Now - LastAdjustment
            Dim diff As Integer = Int(ts.TotalMilliseconds)

            If diff < 100 Then
                LastPass += 0.5
            Else
                LastPass = 1
            End If

            AdjustReadingGuideThickness -= LastPass

            LastAdjustment = Now

            If AdjustReadingGuideThickness < 0 Then AdjustReadingGuideThickness = 0

            SetupForRulerOrReadingGuide()

            ReturnCode = True

        Else

            ReturnCode = False

        End If

        Return ReturnCode

    End Function

    Friend Function FindHeightOfCurrentScreen(ByVal pt As Point) As Integer

        Dim ConvertedPt As Point = New Point(gMinX + pt.X, gMinY + pt.Y)

        For ScreenNumber As Integer = 0 To gScreens.Length - 1
            If gScreens(ScreenNumber).Bounds.Contains(ConvertedPt) Then
                Return gScreens(ScreenNumber).Bounds.Height
            End If
        Next

        Return 0

    End Function

    Private Sub ScrollLeft(ByVal e As System.Windows.Forms.KeyEventArgs)

        Try


            If gLockRight And gLockLeft Then
                Beep()
                Exit Try
            End If

            Dim Increment As Integer = 1
            If e.Alt Then Increment = gBump
            Increment = Increment * MagnifierFactor

            If e.Shift Then

                If NumberingIsLeftToRightHorizontal Then
                    If MeasuringLineX > 1 Then
                        MeasuringLineX -= 1
                        MeasuringLineXComplement = WidthInPixels - MeasuringLineX
                    End If
                Else
                    If MeasuringLineXComplement < NewRuler.Width Then
                        MeasuringLineXComplement += 1
                        MeasuringLineX = WidthInPixels - MeasuringLineXComplement
                    End If
                End If

            ElseIf gLockLeft And RulerIsHorizontal Then

                If WidthInPixels > 100 Then

                    WidthInPixels -= Increment
                    If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                        If WidthInPixels < 100 Then WidthInPixels = 100
                        SetBoxSizes()
                    Else
                        WidthInPixels += Increment
                    End If

                End If

            ElseIf gLockRight And RulerIsHorizontal Then

                WidthInPixels += Increment
                Dim np As Point = New Point(RulerPanel.Location.X - Increment, RulerPanel.Location.Y)
                If ConfirmPurposedMoveIsInsideTheFence(np) Then
                    RulerPanel.Location = np
                    SetBoxSizes()
                Else
                    WidthInPixels -= Increment
                End If

            Else

                Dim np As Point = New Point(RulerPanel.Location.X - Increment, RulerPanel.Location.Y)
                If ConfirmPurposedMoveIsInsideTheFence(np) Then
                    RulerPanel.Location = np
                    MoveFlag = True
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ScrollRight(ByVal e As System.Windows.Forms.KeyEventArgs)

        Try

            If gLockRight And gLockLeft Then
                Beep()
                Exit Try
            End If

            Dim Increment As Integer = 1
            If e.Alt Then Increment = gBump
            Increment = Increment * MagnifierFactor

            If e.Shift Then

                If NumberingIsLeftToRightHorizontal Then
                    If (MeasuringLineX + 1) <= WidthInPixels Then
                        MeasuringLineX += 1
                        MeasuringLineXComplement = WidthInPixels - MeasuringLineX
                    End If
                Else
                    If MeasuringLineXComplement > 1 Then
                        MeasuringLineXComplement -= 1
                        MeasuringLineX = WidthInPixels - MeasuringLineXComplement
                    End If
                End If

            ElseIf gLockRight And RulerIsHorizontal Then

                If WidthInPixels > 100 Then

                    WidthInPixels -= Increment
                    Dim np As Point = New Point(RulerPanel.Location.X + Increment, RulerPanel.Location.Y)
                    If ConfirmPurposedMoveIsInsideTheFence(np) Then
                        If WidthInPixels < 100 Then WidthInPixels = 100
                        RulerPanel.Location = New Point(np)
                        SetBoxSizes()
                    Else
                        WidthInPixels += Increment
                    End If

                End If

            ElseIf gLockLeft And RulerIsHorizontal Then

                WidthInPixels += Increment
                If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                    SetBoxSizes()
                Else
                    WidthInPixels -= Increment
                End If

            Else

                Dim np As Point = New Point(RulerPanel.Location.X + Increment, RulerPanel.Location.Y)
                If ConfirmPurposedMoveIsInsideTheFence(np) Then
                    RulerPanel.Location = New Point(np)
                    MoveFlag = True
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ScrollDown(ByVal e As System.Windows.Forms.KeyEventArgs)

        Try

            If gLockTop And gLockBottom Then
                Beep()
                Exit Try
            End If

            Dim Increment As Integer = 1
            If (e.Alt) Or (e.KeyCode = 32) Then Increment = gBump

            Increment = Increment * MagnifierFactor

            If (e.Shift) And (Not e.KeyCode = 32) Then

                If NumberingIsSmallOnTopToHighOnBottomVertical Then
                    If (MeasuringLineY + 1) <= HeightInPixels Then MeasuringLineY += 1
                Else
                    If (MeasuringLineYComplement > 1) And (MeasuringLineYComplement <= HeightInPixels) Then MeasuringLineY += 1
                End If

                MeasuringLineYComplement = HeightInPixels - MeasuringLineY

            ElseIf gLockBottom And RulerIsVertical Then

                If HeightInPixels > 100 Then

                    HeightInPixels -= Increment
                    If HeightInPixels < 100 Then HeightInPixels = 100

                    Dim np As Point = New Point(RulerPanel.Location.X, RulerPanel.Location.Y + Increment)
                    If ConfirmPurposedMoveIsInsideTheFence(np) Then
                        RulerPanel.Location = np
                        SetBoxSizes()
                    Else
                        HeightInPixels += Increment
                    End If

                End If

            ElseIf gLockTop And RulerIsVertical Then

                HeightInPixels += Increment
                If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                    SetBoxSizes()
                Else
                    HeightInPixels -= Increment
                    If Increment = gBump Then
                        EscortByHeight(1)
                    End If
                End If

            Else

                Dim np As Point = New Point(RulerPanel.Location.X, RulerPanel.Location.Y + Increment)
                If ConfirmPurposedMoveIsInsideTheFence(np) Then
                    RulerPanel.Location = np
                    MoveFlag = True
                Else
                    EscortByLocationY(1)
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ScrollUp(ByVal e As System.Windows.Forms.KeyEventArgs)

        Try

            If gLockTop And gLockBottom Then
                Beep()
                Exit Try
            End If

            Dim Increment As Integer = 1
            If e.Alt Then Increment = gBump
            Increment = Increment * MagnifierFactor

            If e.Shift Then

                If NumberingIsSmallOnTopToHighOnBottomVertical Then
                    If (MeasuringLineY - 1) > 0 Then MeasuringLineY -= 1
                Else
                    If (MeasuringLineYComplement + 1) <= HeightInPixels Then MeasuringLineY -= 1
                End If

                MeasuringLineYComplement = HeightInPixels - MeasuringLineY

            ElseIf gLockTop And RulerIsVertical Then

                If HeightInPixels > 100 Then
                    HeightInPixels -= Increment
                    If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                        If HeightInPixels < 100 Then HeightInPixels = 100
                        SetBoxSizes()
                    Else
                        HeightInPixels += Increment
                    End If

                End If

            ElseIf gLockBottom And RulerIsVertical Then

                HeightInPixels += Increment
                Dim np As Point = New Point(RulerPanel.Location.X, RulerPanel.Location.Y - Increment)
                If ConfirmPurposedMoveIsInsideTheFence(np) Then
                    RulerPanel.Location = np
                    SetBoxSizes()
                Else
                    HeightInPixels -= Increment
                End If

            Else

                Dim np As Point = New Point(RulerPanel.Location.X, RulerPanel.Location.Y - Increment)
                If ConfirmPurposedMoveIsInsideTheFence(np) Then
                    RulerPanel.Location = np
                    MoveFlag = True
                End If

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetupForRulerOrReadingGuide()

        Static ReadingGuideWasShow As Boolean = False

        If gShowReadingGuide Then

            ReadingGuideWasShow = True

            ClearMeasuringLines()

            If RulerIsHorizontal Then
                HeightInPixels = AdjustReadingGuideThickness + 10
            Else
                WidthInPixels = AdjustReadingGuideThickness + 10
            End If

        Else

            If ReadingGuideWasShow Then
                ReadingGuideWasShow = False
                RetoreMeasuringLines()
            End If

            If RulerIsHorizontal Then
                HeightInPixels = 100
            Else
                WidthInPixels = 100
            End If

        End If

        SetBoxSizes()

    End Sub

    Private Sub ClearMeasuringLines()

        ClearRedMeasuringLine()
        HoldMidPointIsOn = MidPointIsOn
        MidPointIsOn = False

    End Sub

    Private Sub RetoreMeasuringLines()

        RestoreRedMeasuringLine()
        MidPointIsOn = HoldMidPointIsOn

    End Sub

    Private Sub ClearRedMeasuringLine()

        HoldRedLineX = MeasuringLineX
        HoldRedLineXComplement = MeasuringLineXComplement
        HoldRedLineY = MeasuringLineY
        HoldRedLineYComplement = MeasuringLineYComplement
        If RulerIsHorizontal Then
            MeasuringLineX = -1
            MeasuringLineXComplement = -1
        Else
            MeasuringLineY = -1
            MeasuringLineYComplement = -1
        End If

    End Sub

    Private Sub RestoreRedMeasuringLine()

        MeasuringLineX = HoldRedLineX
        MeasuringLineXComplement = HoldRedLineXComplement
        MeasuringLineY = HoldRedLineY
        MeasuringLineYComplement = HoldRedLineYComplement

    End Sub

    Private Sub RulerPanel_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RulerPanel.MouseDown

        Try

            MoveFlag = False

            If CheckIn(e.Location, AboutBox) Then
                ShowHelpNow()
                Exit Try
            End If


            If CheckIn(e.Location, HorizontalVeriticalBox) Then
                FlipVerticalHorizontal()
                Exit Try
            End If

            If CheckIn(e.Location, TickPositionBox) Then

                If gShowReadingGuide OrElse (e.Button = Windows.Forms.MouseButtons.Right) Then
                    gShowReadingGuide = Not gShowReadingGuide
                    SetupForRulerOrReadingGuide()

                    ControlsAreHidden = False
                    Timer1.Enabled = True
                    Timer1.Start()
                    RulerDrawBackground()

                Else
                    SwapTicksNow()
                End If

                Exit Try
            End If

            If CheckIn(e.Location, NumberBox) Then
                ShowResizeNow()
                Exit Try
            End If

            If CheckIn(e.Location, BackgroundBox) Then
                ChangeBackgroundNow(True)
                Exit Try
            End If

            If CheckIn(e.Location, MagnifyBox) Then
                MagnifyNow()
                Exit Try
            End If


            If CheckIn(e.Location, CloseBox) Then
                CloseOnDemand = True
                FrmRuler_Close()
                Exit Try
            End If

            If CheckIn(e.Location, MinimizeNowBox) Then
                MinimizeNow()
                Exit Try
            End If

            If gShowReadingGuide Then
            Else
                'Check to see if the user want to draw a red line
                If CheckIn(e.Location, MeasuringBox) Then
                    AbsoluteMousePositionOnRulerX = e.X / MagnifierFactor
                    AbsoluteMousePositionOnRulerY = e.Y / MagnifierFactor

                    RelativeMousePositionOnRulerX = AbsoluteMousePositionOnRulerX
                    RelativeMousePositionOnRulerY = AbsoluteMousePositionOnRulerY

                    If RulerIsHorizontal Then

                        If NumberingIsLeftToRightHorizontal Then
                            MeasuringLineX = RelativeMousePositionOnRulerX + 1
                            MeasuringLineXComplement = WidthInPixels - MeasuringLineX
                        Else
                            MeasuringLineX = RelativeMousePositionOnRulerX
                            MeasuringLineXComplement = WidthInPixels - MeasuringLineX
                        End If

                    Else

                        If NumberingIsSmallOnTopToHighOnBottomVertical Then
                            MeasuringLineY = RelativeMousePositionOnRulerY + 1
                            MeasuringLineYComplement = HeightInPixels - MeasuringLineY
                        Else
                            MeasuringLineY = RelativeMousePositionOnRulerY
                            MeasuringLineYComplement = HeightInPixels - MeasuringLineY
                        End If

                    End If

                    RulerDrawBackground()

                    Exit Try
                End If
            End If

            If e.Button = Windows.Forms.MouseButtons.Left Then

                ' the left mouse button allows ruler to be resized
                If CheckIn(e.Location, ResizeRightOrDownBox) Then
                    ResizeRightFlag = True
                    If RulerIsHorizontal Then
                        Me.Cursor = Cursors.SizeWE
                    Else
                        Me.Cursor = Cursors.SizeNS
                    End If
                    Exit Try
                End If

                If CheckIn(e.Location, ResizeLeftOrUpBox) Then
                    ResizeLeftFlag = True
                    If RulerIsHorizontal Then
                        Me.Cursor = Cursors.SizeWE
                    Else
                        Me.Cursor = Cursors.SizeNS
                    End If
                    Exit Try
                End If

                If CheckIn(e.Location, BackgroundBox) Then
                    ChangeBackgroundNow(False)
                    Exit Try
                End If

            ElseIf e.Button = Windows.Forms.MouseButtons.Right Then

                ' the right mouse button locks ruler
                If CheckIn(e.Location, ResizeRightOrDownBox) Then
                    If RulerIsHorizontal Then
                        gLockRight = Not gLockRight
                    Else
                        gLockBottom = Not gLockBottom
                    End If
                ElseIf CheckIn(e.Location, ResizeLeftOrUpBox) Then
                    If RulerIsHorizontal Then
                        gLockLeft = Not gLockLeft
                    Else
                        gLockTop = Not gLockTop
                    End If
                End If

                SetBoxSizes()

                RulerDrawBackground()

            ElseIf (e.Button = Windows.Forms.MouseButtons.Middle) AndAlso (CheckIn(e.Location, ResizeLeftOrUpBox) OrElse CheckIn(e.Location, ResizeRightOrDownBox)) Then

                ' take a screen capture of all screens

                Dim bmp As Bitmap = New Bitmap(
                        Screen.AllScreens.Sum(Function(s As Screen) s.Bounds.Width),
                        Screen.AllScreens.Sum(Function(s As Screen) s.Bounds.Height))

                Dim gfx As Graphics = Graphics.FromImage(bmp)

                Try

                    gfx.CopyFromScreen(SystemInformation.VirtualScreen.X,
                       SystemInformation.VirtualScreen.Y,
                       0, 0, SystemInformation.VirtualScreen.Size)

                    Dim searchDirection As Integer ' -1 searching to left, 1 searching to right

                    Dim startingX As Integer
                    Dim startingY As Integer
                    Dim endingX As Integer
                    Dim endingY As Integer

                    Dim finalX As Integer = RulerPanel.Location.X
                    Dim finalY As Integer = RulerPanel.Location.Y

                    Dim FoundPixelColourChange As Boolean = False

                    If CheckIn(e.Location, ResizeRightOrDownBox) Then

                        If RulerIsHorizontal Then

                            searchDirection = 1

                            startingX = RulerPanel.Location.X + WidthInPixels + 1
                            startingY = RulerPanel.Location.Y
                            endingX = bmp.Width
                            endingY = startingY

                        Else

                            searchDirection = 1

                            startingX = RulerPanel.Location.X
                            startingY = RulerPanel.Location.Y + HeightInPixels + 1
                            endingX = startingX
                            endingY = bmp.Height

                        End If

                    End If

                    If CheckIn(e.Location, ResizeLeftOrUpBox) Then

                        If RulerIsHorizontal Then

                            searchDirection = -1

                            startingX = RulerPanel.Location.X - 1
                            startingY = RulerPanel.Location.Y
                            endingX = 0
                            endingY = startingY

                        Else

                            searchDirection = -1

                            startingX = RulerPanel.Location.X
                            startingY = RulerPanel.Location.Y - 1
                            endingX = startingX
                            endingY = 0

                        End If

                    End If

                    If RulerHasScaleOnTopOrLeft OrElse gShowReadingGuide Then
                    Else
                        If RulerIsHorizontal Then
                            startingY += 100
                            endingY = startingY
                        Else
                            startingX += 100
                            endingX = startingX
                        End If

                    End If

                    ' search for the first pixel colour change
                    ' if found then calculate the new position and size of the ruler based the pixel just prior to the change
                    ' if not found then calculate the new position and size of the ruler based on the edge of the screen

                    Dim startingPixel As Color = bmp.GetPixel(startingX, startingY)

                    If RulerIsHorizontal Then

                        For x As Integer = startingX To endingX - searchDirection Step searchDirection

                            If startingPixel <> bmp.GetPixel(x, startingY) Then

                                If startingX > x Then
                                    finalX = x + 1
                                    NewSize = WidthInPixels + (startingX - x)
                                Else
                                    NewSize = WidthInPixels + (x - startingX) + 1
                                End If

                                FoundPixelColourChange = True

                                Exit For

                            End If

                        Next

                        If FoundPixelColourChange Then
                        Else
                            ' handles the case we've reach the end of the screen and the pixel colour hasn't changed

                            If searchDirection = -1 Then
                                finalX = 0
                                NewSize = WidthInPixels + startingX + 1
                            Else
                                NewSize = WidthInPixels + endingX - startingX + 1
                            End If

                        End If

                    Else

                        For y As Integer = startingY To endingY - searchDirection Step searchDirection

                            If startingPixel <> bmp.GetPixel(startingX, y) Then

                                If startingY > y Then
                                    finalY = y + 1
                                    NewSize = HeightInPixels + (startingY - y)
                                Else
                                    NewSize = HeightInPixels + (y - startingY) + 1
                                End If

                                FoundPixelColourChange = True

                                Exit For

                            End If

                        Next

                        If FoundPixelColourChange Then
                        Else
                            ' handles the case we've reach the end of the screen and the pixel colour hasn't changed

                            If searchDirection = -1 Then
                                finalY = 0
                                NewSize = HeightInPixels + startingY + 1
                            Else
                                NewSize = HeightInPixels + endingY - startingY + 1
                            End If

                        End If

                    End If

                    ' check that the ruler will still be within the bounds of the screen
                    ' if it is then the reposition the ruler to the new location and resize it to NewSize (as calculated above)
                    ' if it is not then do not reposiiton the ruler and set NewSize to 0 (so it will not be resized)

                    np = New Point(finalX, finalY)
                    If IsThisPointInBounds(np) Then
                        RulerPanel.Location = np
                    Else
                        NewSize = 0
                    End If

                Catch ex As Exception

                End Try

                bmp.Dispose()
                gfx.Dispose()

                Dim hFenceRulerOnScreen As Boolean = gFenceRulerOnScreen
                gFenceRulerOnScreen = True
                ResizeAndFenceAsNeeded()
                gFenceRulerOnScreen = hFenceRulerOnScreen

                Exit Try

            End If

            ' Move the ruler
            RelativeMousePositionOnRulerX = e.X
            RelativeMousePositionOnRulerY = e.Y

            Me.Cursor = Cursors.Hand
            MoveFlag = True

        Catch ex As Exception

            MsgBox(ex.TargetSite.Name & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error") : Me.DialogResult = 0

        End Try

    End Sub

    <System.Diagnostics.DebuggerStepThrough()> Private Function CheckIn(ByRef e As System.Drawing.Point, ByVal box As Rectangle) As Boolean

        ' change in v2.8

        'Dim ReturnValue As Boolean = False

        'If e.X > box.X Then
        '    If e.Y > box.Y Then
        '        If e.X < (box.X + box.Width) Then
        '            If e.Y < (box.Y + box.Height) Then
        '                ReturnValue = True
        '            End If
        '        End If
        '    End If
        'End If

        'Return ReturnValue

        Return (e.X > box.X) AndAlso (e.Y > box.Y) AndAlso (e.X < (box.X + box.Width)) AndAlso (e.Y < (box.Y + box.Height))

    End Function

    Private Sub MinimizeNow()
        Me.WindowState = FormWindowState.Minimized
        frmWhiteBackGround.Hide()
    End Sub

    Private Sub ShowHelpNow()

        Try

            gRefreshRulerAndAboutWindow = True

            While gRefreshRulerAndAboutWindow

                Dim HoldFenceRulerOnScreen As Boolean = gFenceRulerOnScreen

                Do

                    'the help window will need to be reloaded if the language changes
                    gReloadHelp = False

                    'Me.Hide()
                    'Application.DoEvents()

                    Dim frmAbout As Form = New frmAbout
                    frmAbout.ShowDialog()
                    frmAbout.Dispose()

                    'Me.Show()

                Loop Until gReloadHelp = False

                If gShowBalloons Then
                Else
                    ToolTip1.SetToolTip(RulerPanel, "")
                End If
                Me.ToolTip1.Active = gShowBalloons
                frmAbout.Dispose()

                Me.Opacity = gCurrentSkin.Opacity

                MyTransparentColour = gCurrentSkin.TransparencyColour
                Me.TransparencyKey = MyTransparentColour
                Me.BackColor = MyTransparentColour
                Me.Refresh()

                RulerDrawBackground()
                Application.DoEvents()

            End While

        Catch ex As Exception
            MsgBox(ex.TargetSite.Name & vbCrLf & ex.Message, MsgBoxStyle.Critical, gResourceManager.GetString("0067")) : Me.DialogResult = 0
        End Try

    End Sub

    Private Sub SwapTicksNow()

        RulerHasScaleOnTopOrLeft = Not RulerHasScaleOnTopOrLeft
        SetBoxSizes()
        RulerDrawBackground()

    End Sub

    Private Sub ResizeAndFenceAsNeeded()

        If RulerIsHorizontal Then

            If gFenceRulerOnScreen Then

                Dim NewWidth As Integer
                For NewWidth = 100 To NewSize
                    np = New Point(RulerPanel.Location.X + NewWidth - 1, RulerPanel.Location.Y)
                    If IsThisPointInBounds(np) Then
                        WidthInPixels = NewWidth
                    Else
                        Exit For
                    End If
                Next
            Else
                WidthInPixels = NewSize
            End If

            RulerPanel.Width = WidthInPixels

        Else

            If gFenceRulerOnScreen Then

                'escort ruler to its new height

                Dim NewHeight As Integer
                For NewHeight = 100 To NewSize
                    np = New Point(RulerPanel.Location.X, RulerPanel.Location.Y + NewHeight - 1)
                    If IsThisPointInBounds(np) Then
                        HeightInPixels = NewHeight
                    Else
                        Exit For
                    End If
                Next
            Else
                HeightInPixels = NewSize
            End If

            RulerPanel.Height = HeightInPixels

        End If

        SetBoxSizes()
        RulerDrawBackground()

    End Sub
    Private Sub ShowResizeNow()

        Try

            If MagnifierFactor = 1 Then
                gSetLocation = New Point(RulerPanel.Location.X + gMinX, RulerPanel.Location.Y + gMinY)
            Else
                gSetLocation = New Point(RulerPanel.Location.X + gMinX + gPrimaryScreenHome.X, RulerPanel.Location.Y + gMinY + gPrimaryScreenHome.Y)
            End If

            Dim oldSize As Integer

            If RulerIsHorizontal Then
                oldSize = WidthInPixels
                NewSize = WidthInPixels
            Else
                oldSize = HeightInPixels
                NewSize = HeightInPixels
            End If

            Dim frmResize As Form = New frmResize
            frmResize.ShowDialog()
            frmResize.Dispose()
            Application.DoEvents()

            If NewSize = oldSize Then
                RulerDrawBackground()
            Else
                ResizeAndFenceAsNeeded()
            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub FlipVerticalHorizontal()

        Try

            RulerIsHorizontal = Not RulerIsHorizontal ' flip ruler
            RulerIsVertical = Not RulerIsHorizontal ' ensure vertical flag is opposite of horizontal flag
            swap(WidthInPixels, HeightInPixels)

            If gShowReadingGuide Then
                If RulerIsHorizontal Then
                    HeightInPixels = 30
                Else
                    WidthInPixels = 20
                End If
            End If

            FenceRulerFollowingAFlip()

            SetBoxSizes()
            RulerDrawBackground()

        Catch ex As Exception
        End Try

    End Sub

    'Private Function WhichScreenAmIOn() As Integer

    '    Dim ScreenNumber As Integer = 0

    '    For ScreenNumber = 0 To gScreens.Length - 1
    '        If gScreens(ScreenNumber).Bounds.Contains(RulerPanel.Location) Then
    '            Exit For
    '        End If
    '    Next

    '    Return ScreenNumber

    'End Function

    Private Function ConfirmPurposedMoveIsInsideTheFence(ByVal PurposedPoint As Point) As Boolean

        Static Dim ReturnCode As Boolean

        If gFenceRulerOnScreen Then

            ReturnCode = False

            If IsThisPointInBounds(PurposedPoint) Then ' check top left of ruler
                If IsThisPointInBounds(New Point(PurposedPoint.X + WidthInPixels - 1, PurposedPoint.Y)) Then ' check top right of ruler
                    If IsThisPointInBounds(New Point(PurposedPoint.X, PurposedPoint.Y + HeightInPixels - 1)) Then ' check bottom left of ruler
                        If IsThisPointInBounds(New Point(PurposedPoint.X + WidthInPixels - 1, PurposedPoint.Y + HeightInPixels - 1)) Then ' check bottom right of ruler
                            ReturnCode = True
                        End If
                    End If
                End If
            End If

        Else

            ReturnCode = True

        End If

        Return ReturnCode

    End Function

    Private Sub FenceRulerFollowingAFlip()

        If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
        Else

            ' The ruler can not be flipped all the way to the new location
            ' however the following code will escort it to the fence

            If RulerIsHorizontal Then

                If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                Else
                    For WidthInPixels = 100 To WidthInPixels
                        If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                        Else
                            Exit For
                        End If
                    Next
                    WidthInPixels -= 1
                End If

            Else

                If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                Else
                    For HeightInPixels = 100 To HeightInPixels
                        If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                        Else
                            Exit For
                        End If
                    Next
                    HeightInPixels -= 1
                End If

            End If

            RulerDrawAllLines()

        End If

    End Sub

    Private Sub ChangeBackgroundNow(ByVal Advance As Boolean)

        'Last Opacity will be either 0 or 1
        'It will be 1 if the last time thru here the current skin's Opacity was 1
        'It will be 0 if the last time thru here the current skin's Opacity was 0
        'It is initially set to -1 but updated to the Opacity of the current skin on the intial time thru

        Static Dim LastOpacity As Integer = -1
        If LastOpacity = -1 Then LastOpacity = gCurrentSkin.Opacity
        If LastOpacity < 1 Then LastOpacity = 0

        If Advance Then
            If gCurrentSkin.Name = gNextEnabledSkinName Then Exit Sub
        Else
            If gCurrentSkin.Name = gPreviousEnabledSkinName Then Exit Sub
        End If

        Try

            If Advance Then
                gCurrentSkin.Name = gNextEnabledSkinName
            Else
                gCurrentSkin.Name = gPreviousEnabledSkinName
            End If
            Call LoadSkin(gCurrentSkin.Name)

            SetNextAndPreviousEnabledSkinNames()

            Me.Opacity = gCurrentSkin.Opacity
            If Me.Opacity = 1 Then
                frmWhiteBackGround.Hide()
            Else
                If MagnifierFactor > 1 Then
                    frmWhiteBackGround.Show()
                Else
                    frmWhiteBackGround.Hide()
                End If
            End If

            Me.BackColor = gCurrentSkin.TransparencyColour
            Me.TransparencyKey = Me.BackColor

            RulerDrawBackground()

            If MagnifierFactor > 1 Then
                'restore the WhiteBackground
                If gCurrentSkin.Opacity = LastOpacity Then
                Else
                    If gCurrentSkin.Opacity = 1 Then
                        frmWhiteBackGround.Hide()
                    Else
                        PanelLocationX = Me.MagnifyViewPanel.Location.X
                        PanelLocationY = Me.MagnifyViewPanel.Location.Y
                        frmWhiteBackGround.Show()
                    End If
                End If
            End If

            LastOpacity = gCurrentSkin.Opacity

        Catch ex As Exception
            MsgBox(ex.TargetSite.Name & vbCrLf & ex.Message, MsgBoxStyle.Critical, gResourceManager.GetString("0067")) : Me.DialogResult = 0
        End Try

    End Sub

    Private gHoldLocationForMe As Point = New Point(0, 0)
    Private gHoldSizeForMe As Size = Me.Size

    Private Sub MagnifyNow()

        If gMagnifierFound Then

            newMagnifyNow()

        Else

            oldMagnifyNow()

        End If
    End Sub

    Private Sub newMagnifyNow()


        Dim MagniferIsRunning As Boolean = (Diagnostics.Process.GetProcessesByName("magnify").Length > 0)

        If MagniferIsRunning Then

            SetPriority(ProcessPriorityClass.High)
            'close the magnfier
            modKeySend.SendTheTransmissionString("{WIN}{ESC}")
            SetPriority(ProcessPriorityClass.Normal)

        Else

            'start the magnifier
            System.Diagnostics.Process.Start(gMagnifierProgram)
            Threading.Thread.Sleep(250)

        End If

    End Sub

    Private Sub SetPriority(ByVal Prioirty As ProcessPriorityClass)

        Static Dim myProcess As Process = Process.GetCurrentProcess
        myProcess.PriorityClass = Prioirty

    End Sub

    Private Sub oldMagnifyNow()

        ' only the primary screen can be magnified

        Try

            If MagnifierFactor = 1 Then
                gHoldLocationForMe = Me.Location
                gHoldSizeForMe = Me.Size
            End If

            oldMagnifierFactor = MagnifierFactor

            If MagnifierFactor = 1 Then
                gSetLocation = New Point(RulerPanel.Location.X + gMinX, RulerPanel.Location.Y + gMinY)
            Else
                gSetLocation = New Point(RulerPanel.Location.X + gMinX + gPrimaryScreenHome.X, RulerPanel.Location.Y + gMinY + gPrimaryScreenHome.Y)
            End If

            gRulerPanelBounds = RulerPanel.Bounds

            Dim frmMagnify As Form = New frmMagnify
            frmMagnify.ShowDialog()
            frmMagnify.Dispose()

            Me.Activate()

            If oldMagnifierFactor = MagnifierFactor Then
                RulerDrawAllLines()
                Exit Try
            End If

            If MagnifierFactor > 1 Then

                Application.DoEvents()
                System.Threading.Thread.Sleep(500)
                Application.DoEvents()

                ' reset the screen to be orgininal size so a snapshot can be taken of it at its original size 
                If oldMagnifierFactor > 1 Then
                    Dim hold As Integer = MagnifierFactor
                    RestoreNormalScreen()
                    MagnifierFactor = hold
                End If

                'Hide everything before taking a snapshot
                RulerPanel.Visible = False
                frmWhiteBackGround.Hide()
                MagnifyViewPanel.Visible = False
                While RulerPanel.Visible OrElse frmWhiteBackGround.Visible OrElse MagnifyViewPanel.Visible
                    Application.DoEvents()
                    System.Threading.Thread.Sleep(100)
                End While

                'take the snapshot
                Dim sc As New ScreenShotLogic.ScreenCapture
                ScreenCapture = New Bitmap(sc.CaptureScreen.GetThumbnailImage(Screen.PrimaryScreen.Bounds.Width * MagnifierFactor, Screen.PrimaryScreen.Bounds.Height * MagnifierFactor, Nothing, Nothing))
                sc = Nothing

                'restore the ruler
                RulerPanel.Visible = True

                'Show the snapshot of the primary screen in an scollable window on the primary screen
                Me.Location = New Point(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y)
                Me.TransparencyKey = MyTransparentColour
                Me.Size = New Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
                Me.AutoScroll = True

                MagnifyViewPanel.Width = ScreenCapture.Width
                MagnifyViewPanel.Height = ScreenCapture.Height
                MagnifyViewPanel.BackgroundImageLayout = ImageLayout.None
                MagnifyViewPanel.BackgroundImage = ScreenCapture
                MagnifyViewPanel.Visible = True

                PanelLocationX = Me.MagnifyViewPanel.Location.X
                PanelLocationY = Me.MagnifyViewPanel.Location.Y

                If gCurrentSkin.Opacity < 1 Then
                    frmWhiteBackGround.Show()
                End If

                RulerPanel.Location = New Point((gMinX + RulerPanel.Location.X) * MagnifierFactor, (gMinY + RulerPanel.Location.Y) * MagnifierFactor)
                SetBoxSizes()

                Me.PilotFishPanel.Location = New Point(Me.RulerPanel.Location.X + 400, Me.RulerPanel.Location.Y + 400)
                Me.ScrollControlIntoView(Me.PilotFishPanel)

                RulerPanel.Visible = True

                Me.Size = New Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
                Me.AutoScroll = True

            Else

                MagnifierFactor = oldMagnifierFactor
                RestoreNormalScreen()

            End If

            Me.BringToFront()
            RulerDrawBackground()
            Application.DoEvents()

            Me.Activate()
            Flush()

        Catch ex As Exception
            MsgBox(ex.TargetSite.Name & vbCrLf & ex.Message, MsgBoxStyle.Critical, gResourceManager.GetString("0067")) : Me.DialogResult = 0
        End Try

    End Sub

    Public Sub RestoreNormalScreen_Public()
        Call Me.BeginInvoke(Me.CallRestoreNormalScreen)
    End Sub
    Dim CallRestoreNormalScreen As New MethodInvoker(AddressOf Me.RestoreNormalScreen)

    Private Sub RestoreNormalScreen()

        Try

            Me.Location = gHoldLocationForMe
            Me.Size = gHoldSizeForMe
            Me.AutoScroll = False

            RulerDrawBackground_InProgress = True

            Me.BackColor = MyTransparentColour
            Me.TransparencyKey = MyTransparentColour

            MagnifyViewPanel.BackgroundImageLayout = ImageLayout.Stretch

            If MagnifyViewPanel.BackgroundImage IsNot Nothing Then
                MagnifyViewPanel.BackgroundImage.Dispose()  ' v3.3 testing prefer this line
            End If

            MagnifyViewPanel.Visible = False

            frmWhiteBackGround.Hide()

            RulerPanel.Visible = True
            RulerPanel.Location = New Point(RulerPanel.Location.X / MagnifierFactor - gMinX, RulerPanel.Location.Y / MagnifierFactor - gMinY)

            oldMagnifierFactor = 1
            MagnifierFactor = 1

            Me.BringToFront()
            SetBoxSizes()

            RulerDrawBackground_InProgress = False

            RulerDrawBackground()
            frmWhiteBackGround.Hide()
            RulerPanel.BringToFront()
            Application.DoEvents()

        Catch ex As Exception
            MsgBox(ex.TargetSite.Name & vbCrLf & ex.Message, MsgBoxStyle.Critical, gResourceManager.GetString("0067")) : Me.DialogResult = 0
        End Try

    End Sub

    Private Sub swap(ByRef a As Integer, ByRef b As Integer)
        Dim h As Integer = a : a = b : b = h
    End Sub

    Private Sub RulerPanel_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RulerPanel.MouseUp

        If MoveFlag OrElse ResizeRightFlag OrElse ResizeLeftFlag Then
            MoveFlag = False
            ResizeRightFlag = False
            ResizeLeftFlag = False
            SetBoxSizes()
            RulerDrawBackground()
        Else
            MoveFlag = False
            ResizeRightFlag = False
            ResizeLeftFlag = False
        End If

        Me.Cursor = Cursors.Default

        Flush()

    End Sub

    Private Sub RulerPanel_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RulerPanel.MouseMove

        Static Dim MoveUnderway As Boolean = False
        If MoveUnderway Then Exit Sub

        If e.Button = Windows.Forms.MouseButtons.Left OrElse e.Button = Windows.Forms.MouseButtons.Right Then
        Else
            Exit Sub
        End If

        MoveUnderway = True
        Try

            'prevent ruler from being resized on any side that is locked
            If RulerIsHorizontal Then
                If ResizeRightFlag And gLockRight Then Exit Try
                If ResizeLeftFlag And gLockLeft Then Exit Try
            Else
                If ResizeRightFlag And gLockBottom Then Exit Try
                If ResizeLeftFlag And gLockTop Then Exit Try
            End If

            'resize ruler by dragging it

            Dim oldWidthInPixels As Integer = WidthInPixels
            Dim oldHeightInPixels As Integer = HeightInPixels

            If ResizeRightFlag Then

                ' don't let ruler get smaller then 100 pixel (height/width)
                If RulerIsHorizontal Then

                    WidthInPixels = ((e.X + 1) / MagnifierFactor)

                    If WidthInPixels < 100 Then
                        WidthInPixels = 100
                    End If

                    If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                    Else

                        Dim NewWidthInPixels As Integer = WidthInPixels
                        WidthInPixels = oldWidthInPixels
                        If oldWidthInPixels < NewWidthInPixels Then
                            EscortByWidth(1)
                        ElseIf oldWidthInPixels > NewWidthInPixels Then
                            EscortByWidth(-1)
                        End If
                    End If

                Else

                    HeightInPixels = ((e.Y + 1) / MagnifierFactor)
                    If HeightInPixels < 100 Then
                        HeightInPixels = 100
                    End If

                    If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                    Else

                        Dim NewHeightInPixels As Integer = HeightInPixels
                        HeightInPixels = oldHeightInPixels
                        If oldHeightInPixels < NewHeightInPixels Then
                            EscortByHeight(1)
                        ElseIf oldHeightInPixels > NewHeightInPixels Then
                            EscortByHeight(-1)
                        End If

                    End If

                End If

                SetBoxSizes()
                RulerDrawBackground()

                Exit Try

            End If

            If ResizeLeftFlag Then

                If e.Y = 0 Then Exit Try

                ' don't let ruler get smaller then 100 pixel (height/width)
                If RulerIsHorizontal Then
                    Dim FixedX As Integer = RulerPanel.Location.X + WidthInPixels - 100
                    WidthInPixels -= e.X / MagnifierFactor
                    If WidthInPixels < 100 Then
                        WidthInPixels = 100
                        RulerPanel.Location = New Point(FixedX, RulerPanel.Location.Y)
                    Else
                        Dim np As Point = New Point(RulerPanel.Location.X + e.X, RulerPanel.Location.Y)
                        If ConfirmPurposedMoveIsInsideTheFence(np) Then
                            RulerPanel.Location = np
                        Else

                            If oldWidthInPixels < WidthInPixels Then
                                EscortByWidth(1)
                            ElseIf oldWidthInPixels > WidthInPixels Then
                                EscortByWidth(-1)
                            End If

                        End If

                    End If

                Else

                    Dim FixedY As Integer = RulerPanel.Location.Y + HeightInPixels - 100
                    HeightInPixels -= e.Y / MagnifierFactor
                    If HeightInPixels < 100 Then
                        HeightInPixels = 100
                        RulerPanel.Location = New Point(RulerPanel.Location.X, FixedY)
                    Else

                        Dim np As Point = New Point(RulerPanel.Location.X, RulerPanel.Location.Y + e.Y)
                        If ConfirmPurposedMoveIsInsideTheFence(np) Then
                            RulerPanel.Location = np
                        Else

                            If oldHeightInPixels < HeightInPixels Then
                                EscortByHeight(1)
                            ElseIf oldHeightInPixels > HeightInPixels Then
                                EscortByHeight(-1)
                            End If

                        End If
                    End If
                End If
                SetBoxSizes()
                RulerDrawBackground()
                Exit Try

            End If

            If MoveFlag Then

                np = New Point(RulerPanel.Location.X + e.X - RelativeMousePositionOnRulerX, RulerPanel.Location.Y + e.Y - RelativeMousePositionOnRulerY)

                If ConfirmPurposedMoveIsInsideTheFence(np) Then

                    If RulerPanel.Location <> np Then
                        MoveRulerToANewLocation(np)
                        RulerDrawAllLines()
                    End If

                Else

                    ' The ruler can not be moved all the way to the new location
                    ' however the following code will escort it to the fence

                    Dim DeltaX As Integer = np.X - RulerPanel.Location.X
                    Dim DeltaY As Integer = np.Y - RulerPanel.Location.Y

                    'Find out if X , Y or both were out of bounds
                    Dim X_OutOfBounds As Boolean = Not ConfirmPurposedMoveIsInsideTheFence(New Point(RulerPanel.Location.X + e.X - RelativeMousePositionOnRulerX, RulerPanel.Location.Y))
                    Dim Y_OutOfBounds As Boolean = Not ConfirmPurposedMoveIsInsideTheFence(New Point(RulerPanel.Location.X, RulerPanel.Location.Y + e.Y - RelativeMousePositionOnRulerY))

                    If X_OutOfBounds Then
                        If DeltaX > 0 Then
                            EscortByLocationX(1)
                        ElseIf DeltaX < 0 Then
                            EscortByLocationX(-1)
                        End If
                    End If

                    If Y_OutOfBounds Then

                        If DeltaY > 0 Then
                            EscortByLocationY(1)
                        ElseIf DeltaY < 0 Then
                            EscortByLocationY(-1)
                        End If

                    End If

                    RulerDrawAllLines()

                End If

            End If

        Catch ex As Exception
            MsgBox(ex.TargetSite.Name & vbCrLf & ex.Message, MsgBoxStyle.Critical, gResourceManager.GetString("0067")) : Me.DialogResult = 0
        End Try

        MoveUnderway = False

    End Sub

    Private Sub EscortByLocationX(ByVal Increment As Integer)

        np = RulerPanel.Location
        For x As Integer = 1 To 250
            np = New Point(np.X + Increment, np.Y)
            If ConfirmPurposedMoveIsInsideTheFence(np) Then
            Else
                np = New Point(np.X - Increment, np.Y)
                Exit For
            End If
        Next
        RulerPanel.Location = np

    End Sub

    Private Sub EscortByLocationY(ByVal Increment As Integer)

        np = RulerPanel.Location
        For x As Integer = 1 To 250
            np = New Point(np.X, np.Y + Increment)
            If ConfirmPurposedMoveIsInsideTheFence(np) Then
            Else
                np = New Point(np.X, np.Y - Increment)
                Exit For
            End If
        Next
        RulerPanel.Location = np

    End Sub

    Private Sub EscortByWidth(ByVal Increment As Integer)

        For x As Integer = 1 To 250
            If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                WidthInPixels += Increment
            Else
                WidthInPixels -= Increment
                Exit For
            End If
        Next

    End Sub

    Private Sub EscortByHeight(ByVal Increment As Integer)

        For x As Integer = 1 To 250
            If ConfirmPurposedMoveIsInsideTheFence(RulerPanel.Location) Then
                HeightInPixels += Increment
            Else
                HeightInPixels -= Increment
                Exit For
            End If
        Next

    End Sub

    Private Sub MoveRulerToANewLocation(ByVal NewLocation As Point)

        If RulerIsHorizontal Then

            If gLockLeft Or gLockRight Then
                RulerPanel.Location = New Point(RulerPanel.Location.X, NewLocation.Y)
            Else
                RulerPanel.Location = NewLocation
            End If
        Else
            If gLockTop Or gLockBottom Then
                RulerPanel.Location = New Point(NewLocation.X, RulerPanel.Location.Y)
            Else
                RulerPanel.Location = NewLocation
            End If
        End If

    End Sub

    Dim MeWindowState As System.Windows.Forms.FormWindowState
    Dim MePaintUnderway As Boolean = False

    Private Sub Me_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        If MePaintUnderway Then Exit Sub

        If MagnifierFactor = 1 AndAlso MoveFlag Then Exit Sub


        MePaintUnderway = True

        MeWindowState = Me.WindowState

        If MeWindowState = FormWindowState.Normal Then

            If MagnifierFactor > 1 Then

                If gCurrentSkin.Opacity < 1 Then 'v3.3

                    frmWhiteBackGround.Show()

                End If

            End If

            Me.Size = New Size(gMaxWidth, gMaxHeight)

        ElseIf MeWindowState = FormWindowState.Minimized Then
            frmWhiteBackGround.Hide()

        End If

        MePaintUnderway = False

    End Sub

    Friend RulerDrawBackground_InProgress As Boolean = False

    Private Sub RulerDrawBackground()

        Try

            If RulerDrawBackground_InProgress Then Exit Try

            If MoveFlag AndAlso (MagnifierFactor = 1) Then Exit Sub ' no need to redraw ruler if only location is changing

            RulerDrawBackground_InProgress = True

            Static Dim w As Integer = -1
            Static Dim h As Integer = -1

            Dim HeightOrWidthChanged As Boolean = ((WidthInPixels <> w) Or (HeightInPixels <> h))

            If HeightOrWidthChanged Then

                If RulerIsHorizontal Then
                    RulerPanel.Size = New Size(WidthInPixels * MagnifierFactor, HeightInPixels)
                Else
                    RulerPanel.Size = New Size(WidthInPixels, HeightInPixels * MagnifierFactor)
                End If

            End If

            RulerPanel.BackgroundImage = Nothing 'v3.3 Do not change this line

            If gCurrentSkin.Tiled Then
                If RulerPanel.BackgroundImageLayout <> ImageLayout.Tile Then
                    RulerPanel.BackgroundImageLayout = ImageLayout.Tile
                End If
            Else
                If RulerPanel.BackgroundImageLayout <> ImageLayout.Stretch Then
                    RulerPanel.BackgroundImageLayout = ImageLayout.Stretch
                End If
            End If

            If MagnifierFactor > 1 Then Application.DoEvents()

            If RulerIsHorizontal Then
                RulerPanel.BackgroundImage = gCurrentSkin.HorizontalBackground
            Else
                RulerPanel.BackgroundImage = gCurrentSkin.VerticalBackground
            End If

            If MagnifierFactor > 1 Then Application.DoEvents()

        Catch ex As Exception

        End Try

        RulerDrawBackground_InProgress = False

    End Sub

    Private Sub SetBoxSizes()

        Dim WidthInUnits, HeightInUnits As Integer

        If RulerIsHorizontal Then
            WidthInUnits = WidthInPixels * MagnifierFactor
            HeightInUnits = HeightInPixels
        Else
            WidthInUnits = WidthInPixels
            HeightInUnits = HeightInPixels * MagnifierFactor
        End If

        Static Dim A1, B1, C1, D1, E1, F1, A2, B2, C2, D2, E2, F2 As Rectangle
        Static Dim A3, B3, C3, D3, E3, F3, A4, B4, C4, D4, E4, F4 As Rectangle
        ' A1  B1    A2  B2
        ' C1  D1    C2  D2
        ' E1  F1    E2  F2
        ' 
        ' A3  B3    A4  B4
        ' C3  D3    C4  D4
        ' E3  F3    E4  F4


        '*********************

        A1 = New Rectangle(0, 0, gBoxScale, gBoxScale)
        B1 = New Rectangle(A1.X + gBoxScale, A1.Y, gBoxScale, gBoxScale)
        C1 = New Rectangle(A1.X, A1.Y + gBoxScale, gBoxScale, gBoxScale)
        D1 = New Rectangle(A1.X + gBoxScale, A1.Y + gBoxScale, gBoxScale, gBoxScale)
        E1 = New Rectangle(A1.X, A1.Y + 2 * gBoxScale, gBoxScale, gBoxScale)
        F1 = New Rectangle(A1.X + gBoxScale, A1.Y + 2 * gBoxScale, gBoxScale, gBoxScale)

        A2 = New Rectangle(WidthInUnits - 2 * gBoxScale, 0, gBoxScale, gBoxScale)
        B2 = New Rectangle(A2.X + gBoxScale, A2.Y, gBoxScale, gBoxScale)
        C2 = New Rectangle(A2.X, A2.Y + gBoxScale, gBoxScale, gBoxScale)
        D2 = New Rectangle(A2.X + gBoxScale, A2.Y + gBoxScale, gBoxScale, gBoxScale)
        E2 = New Rectangle(A2.X, A2.Y + 2 * gBoxScale, gBoxScale, gBoxScale)
        F2 = New Rectangle(A2.X + gBoxScale, A2.Y + 2 * gBoxScale, gBoxScale, gBoxScale)

        A3 = New Rectangle(0, HeightInUnits - 3 * gBoxScale, gBoxScale, gBoxScale)
        B3 = New Rectangle(A3.X + gBoxScale, A3.Y, gBoxScale, gBoxScale)
        C3 = New Rectangle(A3.X, A3.Y + gBoxScale, gBoxScale, gBoxScale)
        D3 = New Rectangle(A3.X + gBoxScale, A3.Y + gBoxScale, gBoxScale, gBoxScale)
        E3 = New Rectangle(A3.X, A3.Y + 2 * gBoxScale, gBoxScale, gBoxScale)
        F3 = New Rectangle(A3.X + gBoxScale, A3.Y + 2 * gBoxScale, gBoxScale, gBoxScale)

        A4 = New Rectangle(WidthInUnits - 2 * gBoxScale, HeightInUnits - 3 * gBoxScale, gBoxScale, gBoxScale)
        B4 = New Rectangle(A4.X + gBoxScale, A4.Y, gBoxScale, gBoxScale)
        C4 = New Rectangle(A4.X, A4.Y + gBoxScale, gBoxScale, gBoxScale)
        D4 = New Rectangle(A4.X + gBoxScale, A4.Y + gBoxScale, gBoxScale, gBoxScale)
        E4 = New Rectangle(A4.X, A4.Y + 2 * gBoxScale, gBoxScale, gBoxScale)
        F4 = New Rectangle(A4.X + gBoxScale, A4.Y + 2 * gBoxScale, gBoxScale, gBoxScale)

        If Not MoveFlag Then

            If RulerIsHorizontal Then

                If RulerHasScaleOnTopOrLeft Or gShowReadingGuide Then

                    MagnifyBox = A3
                    BackgroundBox = B3
                    ResizeLeftOrUpBox = C3
                    NumberBox = D3
                    HorizontalVeriticalBox = E3
                    TickPositionBox = F3

                    AboutBox = C4
                    ResizeRightOrDownBox = D4
                    MinimizeNowBox = E4
                    CloseBox = F4

                    MasterBoxA = New Rectangle(A3.X, A3.Y, 2 * gBoxScale, 3 * gBoxScale)
                    MasterBoxB = New Rectangle(C4.X, C4.Y, 2 * gBoxScale, 2 * gBoxScale)

                Else

                    HorizontalVeriticalBox = A1
                    TickPositionBox = B1
                    ResizeLeftOrUpBox = C1
                    NumberBox = D1
                    MagnifyBox = E1
                    BackgroundBox = F1

                    MinimizeNowBox = A2
                    CloseBox = B2
                    AboutBox = C2
                    ResizeRightOrDownBox = D2

                    MasterBoxA = New Rectangle(A1.X, A1.Y, 2 * gBoxScale, 3 * gBoxScale)
                    MasterBoxB = New Rectangle(A2.X, A2.Y, 2 * gBoxScale, 2 * gBoxScale)

                End If

            Else

                If RulerHasScaleOnTopOrLeft Or gShowReadingGuide Then

                    ResizeLeftOrUpBox = A2
                    HorizontalVeriticalBox = B2
                    NumberBox = C2
                    TickPositionBox = D2
                    BackgroundBox = E2
                    MagnifyBox = F2

                    AboutBox = C4
                    MinimizeNowBox = D4
                    ResizeRightOrDownBox = E4
                    CloseBox = F4

                    MasterBoxA = New Rectangle(A2.X, A2.Y, 2 * gBoxScale, 3 * gBoxScale)
                    MasterBoxB = New Rectangle(C4.X, C4.Y, 2 * gBoxScale, 2 * gBoxScale)

                Else

                    HorizontalVeriticalBox = A1
                    ResizeLeftOrUpBox = B1
                    TickPositionBox = C1
                    NumberBox = D1
                    MagnifyBox = E1
                    BackgroundBox = F1

                    MinimizeNowBox = C3
                    AboutBox = D3
                    CloseBox = E3
                    ResizeRightOrDownBox = F3

                    MasterBoxA = New Rectangle(A1.X, A1.Y, 2 * gBoxScale, 3 * gBoxScale)
                    MasterBoxB = New Rectangle(C3.X, C3.Y, 2 * gBoxScale, 2 * gBoxScale)

                End If

            End If

        End If

    End Sub

    Private Sub RulerPanel_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles RulerPanel.Paint

        RulerDrawAllLines()

    End Sub

    Private Sub RulerDrawAllLines()

        If RulerIsHorizontal Then
            DrawRuler(WidthInPixels * MagnifierFactor, HeightInPixels)
        Else
            DrawRuler(WidthInPixels, HeightInPixels * MagnifierFactor)
        End If

    End Sub

    Private lHoldThickness As Integer = 0
    Private Sub DrawRuler(ByVal WidthInUnits As Integer, ByVal HeightInUnits As Integer)

        Try

            g = RulerPanel.CreateGraphics ' changed in v2.4.5

            NewRuler = New Rectangle(0, 0, WidthInUnits, HeightInUnits)

            If MagnifierFactor > 1 Then SetBoxSizes()

            'v3.4.3
            If gShowReadingGuide Then

                If ControlsAreHidden Then

                Else

                    If RulerIsHorizontal Then

                        If gAutoExpand Then

                            If AdjustReadingGuideThickness < 50 Then
                                lHoldThickness = HeightInPixels
                                AdjustReadingGuideThickness = 50
                                SetupForRulerOrReadingGuide()
                                RulerDrawBackground()
                            End If

                        Else

                            If AdjustReadingGuideThickness < 21 Then
                                lHoldThickness = HeightInPixels
                                AdjustReadingGuideThickness = 21
                                SetupForRulerOrReadingGuide()
                                RulerDrawBackground()
                            End If

                        End If

                    Else

                        If AdjustReadingGuideThickness < 10 Then
                            lHoldThickness = WidthInPixels
                            AdjustReadingGuideThickness = 10
                            SetupForRulerOrReadingGuide()
                            RulerDrawBackground()
                        End If

                    End If

                End If

            End If

            Dim MeasuringBoxHeight As Integer = 45

            If RulerIsHorizontal Then
                If RulerHasScaleOnTopOrLeft Then
                    MeasuringBox = New Rectangle(NewRuler.X, NewRuler.Y, NewRuler.Width, MeasuringBoxHeight)
                Else
                    MeasuringBox = New Rectangle(NewRuler.X, NewRuler.Y + NewRuler.Height - MeasuringBoxHeight, NewRuler.Width, MeasuringBoxHeight)
                End If
            Else
                If RulerHasScaleOnTopOrLeft Then
                    MeasuringBox = New Rectangle(NewRuler.X, NewRuler.Y, MeasuringBoxHeight, NewRuler.Height)
                Else
                    MeasuringBox = New Rectangle(WidthInPixels - MeasuringBoxHeight, NewRuler.Y, WidthInPixels, NewRuler.Height)
                End If
            End If

            If gShowReadingGuide AndAlso ControlsAreHidden Then
            Else

                DrawAQuestionMarkInABox(g, AboutBox)
                DrawAPlusInABox(g, MagnifyBox)
                DrawANumberSignBox(g, NumberBox)
                DrawASquareInABox(g, ResizeLeftOrUpBox)
                DrawASlashInABox(g, HorizontalVeriticalBox)
                DrawTicksInABox(g, TickPositionBox, RulerIsHorizontal, RulerHasScaleOnTopOrLeft)
                DrawACircleInABox(g, BackgroundBox)
                DrawASquareInABox(g, ResizeRightOrDownBox)
                DrawAMinusInABox(g, MinimizeNowBox)
                DrawAnXInABox(g, CloseBox)

            End If

            If RulerIsHorizontal Then
                If gLockLeft Then DrawALockedBox(g, ResizeLeftOrUpBox)
                If gLockRight Then DrawALockedBox(g, ResizeRightOrDownBox)
            Else
                If gLockTop Then DrawALockedBox(g, ResizeLeftOrUpBox)
                If gLockBottom Then DrawALockedBox(g, ResizeRightOrDownBox)
            End If


            'draw edges of ruler

            If gShowReadingGuide AndAlso gNoLines Then

            Else



                If RulerIsHorizontal Then

                    If RulerHasScaleOnTopOrLeft Then
                        g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), 0, HeightInUnits - 1, WidthInUnits, HeightInUnits - 1)
                        If gShowReadingGuide Then g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), 0, 0, WidthInUnits, 0)
                    Else
                        g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), 0, 0, WidthInUnits, 0)
                        If gShowReadingGuide Then g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), 0, HeightInUnits - 1, WidthInUnits, HeightInUnits - 1)
                    End If

                    g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), 0, 0, 0, HeightInUnits) ' left side
                    g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), WidthInUnits - 1, 0, WidthInUnits - 1, HeightInUnits) ' right side

                Else

                    g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), 0, 0, WidthInUnits, 0) ' top
                    g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), 0, HeightInUnits - 1, WidthInUnits, HeightInUnits - 1) ' bottom

                    If RulerHasScaleOnTopOrLeft Then
                        If gShowReadingGuide Then g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), 0, 0, 0, HeightInUnits) ' left side
                        g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), WidthInUnits - 1, 0, WidthInUnits - 1, HeightInUnits) ' right side
                    Else
                        g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), 0, 0, 0, HeightInUnits) ' left side
                        If gShowReadingGuide Then g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), WidthInUnits - 1, 0, WidthInUnits - 1, HeightInUnits) ' right side
                    End If

                End If

            End If

            'show ticks, numbers, and measuring lines
            If gShowReadingGuide Then
            Else
                MePaintUnderway = True
                ShowTicksAndNumbers()
                DrawMeasuringLines()
                MePaintUnderway = False
            End If

        Catch ex As Exception
            MsgBox(ex.TargetSite.Name & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error") : Me.DialogResult = 0
        End Try

    End Sub

    Private Sub CalculateMidPoint()

        If RulerIsHorizontal Then

            MidpointLineX = DividByTwoAndRoundUp(WidthInPixels)
            MidpointLineXComplement = MidpointLineX

        Else

            MidpointLineY = DividByTwoAndRoundUp(HeightInPixels)
            MidpointLineYComplement = MidpointLineY

        End If

    End Sub

    Private Function DividByTwoAndRoundUp(ByVal iNumber As Integer) As Integer

        Dim ReturnValue As Integer = 0

        Select Case Microsoft.VisualBasic.Right(CType(iNumber, String), 1)
            Case Is = "1", "3", "5", "7", "9"
                ReturnValue = Math.Truncate(iNumber / 2) + 1
            Case Else
                ReturnValue = Math.Truncate(iNumber / 2)
        End Select

        'Console.WriteLine("iNumber {0} ReturnValue {1}", iNumber, ReturnValue)
        Return ReturnValue

    End Function

    Private Sub ShowTicksAndNumbers()

        Static Dim Increment, NumbersWritenAt As Integer

        Dim z1 As Integer = 0
        Dim z2 As Integer = 0
        Dim z3 As Integer = 0
        Dim y1 As Integer

        ' Draw Ticks

        Dim TickLocations(3) As Integer
        Dim pixelTickLocations() As Integer = {5, 10, 50, 100}
        Dim emTickLocations() As Integer = {8, 16, 32, 64}
        Dim Custom() As Integer = {0, 0, 0, 0}

        If gRulerScalingFactorSetByUser = CSng(0.0625) Then

            Array.Copy(emTickLocations, TickLocations, 4)

        ElseIf (gRulerScalingFactorSetByUser < CSng(0.05)) OrElse (gRulerScalingFactorSetByUser > CSng(5)) Then

            g.DrawString("Invalid scale.", NumbersFont, New SolidBrush(gCurrentSkin.Numbers), 40, 10)
            g.DrawString("Scale must be between 5.0 and 0.5 inclusive, or", NumbersFont, New SolidBrush(gCurrentSkin.Numbers), 40, 30)
            g.DrawString("be 0.0625 (used for ems).", NumbersFont, New SolidBrush(gCurrentSkin.Numbers), 40, 50)
            Exit Sub

        Else

            Dim factor As Single = CSng(gRulerScalingFactorSetByUser * 100 / CSng(20))
            Custom(0) = factor
            Custom(1) = Custom(0) * 2
            Custom(2) = Custom(1) * 5
            Custom(3) = Custom(2) * 2
            Array.Copy(Custom, TickLocations, 4)

        End If

        Dim BarHeight As Integer

        If RulerHasScaleOnTopOrLeft Then
            BarHeight = 15
        Else
            BarHeight = -15
        End If

        If RulerIsVertical Then
            If HeightInPixels > 1015 Then
                If RulerHasScaleOnTopOrLeft Then
                    BarHeight = 12
                Else
                    BarHeight = -12
                End If
            End If
        End If

        If RulerHasScaleOnTopOrLeft Then
            y1 = 0
            Increment = 20
            NumbersWritenAt = 40
        Else
            y1 = HeightInPixels
            Increment = -20
            NumbersWritenAt = +36
        End If

        Dim StartAt, EndAt, Direction As Single
        If RulerIsHorizontal Then
            If NumberingIsLeftToRightHorizontal Then
                Direction = 1
            Else
                Direction = -1
            End If
        Else
            If NumberingIsSmallOnTopToHighOnBottomVertical Then
                Direction = 1
            Else
                Direction = -1
            End If
        End If

        For Depth As Short = 0 To 3

            Dim StepValue As Integer = TickLocations(Depth) * Direction

            If RulerIsHorizontal Then

                If NumberingIsLeftToRightHorizontal Then
                    StartAt = -1
                    EndAt = WidthInPixels - 1
                Else
                    StartAt = WidthInPixels
                    EndAt = -1
                End If

                For x As Single = StartAt To EndAt Step StepValue

                    If Depth > 0 Then
                        If x = 0 Then z1 = 0 Else z1 = -1
                        If x = WidthInPixels Then z2 = 0 Else z2 = 1
                    End If

                    If MagnifierFactor > 1 Then
                        For z3 = 0 To MagnifierFactor - 1
                            g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), x * MagnifierFactor + z3, y1, x * MagnifierFactor + z3, y1 + BarHeight)
                        Next
                        If Depth > 0 Then
                            g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), x * MagnifierFactor - 1, y1, x * MagnifierFactor - 1, y1 + BarHeight)
                            g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), x * MagnifierFactor + MagnifierFactor, y1, x * MagnifierFactor + MagnifierFactor, y1 + BarHeight)
                        End If
                    Else
                        For z As Integer = z1 To z2
                            For z3 = 0 To MagnifierFactor - 1
                                g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), (x + z) * MagnifierFactor + z3, y1, (x + z) * MagnifierFactor + z3, y1 + BarHeight)
                            Next
                        Next
                    End If

                Next

            Else

                If NumberingIsSmallOnTopToHighOnBottomVertical Then
                    StartAt = -1
                    EndAt = HeightInPixels - 1
                Else
                    StartAt = HeightInPixels
                    EndAt = 0
                End If

                If RulerHasScaleOnTopOrLeft Then

                    For x As Single = StartAt To EndAt Step StepValue

                        If Depth > 0 Then
                            If x = 0 Then z1 = 0 Else z1 = -1
                            If x = HeightInPixels Then z2 = 0 Else z2 = 1
                        End If
                        If MagnifierFactor > 1 Then
                            For z3 = 0 To MagnifierFactor - 1
                                g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), y1, x * MagnifierFactor + z3, y1 + BarHeight, x * MagnifierFactor + z3)
                            Next
                            If Depth > 0 Then
                                g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), y1, x * MagnifierFactor + MagnifierFactor, y1 + BarHeight, x * MagnifierFactor + MagnifierFactor)
                                g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), y1, x * MagnifierFactor - 1, y1 + BarHeight, x * MagnifierFactor - 1)
                            End If
                        Else
                            For z As Integer = z1 To z2
                                For z3 = 0 To MagnifierFactor - 1
                                    g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), y1, (x + z) * MagnifierFactor + z3, y1 + BarHeight, (x + z) * MagnifierFactor + z3)
                                Next
                            Next
                        End If

                    Next
                Else

                    For x As Single = StartAt To EndAt Step StepValue

                        If Depth > 0 Then
                            If x = 0 Then z1 = 0 Else z1 = -1
                            If x = HeightInPixels Then z2 = 0 Else z2 = 1
                        End If
                        If MagnifierFactor > 1 Then
                            For z3 = 0 To MagnifierFactor - 1
                                g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), WidthInPixels + (Depth + 1) * BarHeight, x * MagnifierFactor + z3, WidthInPixels + Depth * BarHeight, x * MagnifierFactor + z3)
                            Next
                            If Depth > 0 Then
                                g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), WidthInPixels + (Depth + 1) * BarHeight, x * MagnifierFactor + MagnifierFactor, WidthInPixels + Depth * BarHeight, x * MagnifierFactor + MagnifierFactor)
                                g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), WidthInPixels + (Depth + 1) * BarHeight, x * MagnifierFactor - 1, WidthInPixels + Depth * BarHeight, x * MagnifierFactor - 1)
                            End If
                        Else
                            For z As Integer = z1 To z2
                                For z3 = 0 To MagnifierFactor - 1
                                    g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), WidthInPixels + (Depth + 1) * BarHeight, (x + z) * MagnifierFactor + z3, WidthInPixels + Depth * BarHeight, (x + z) * MagnifierFactor + z3)
                                Next
                            Next
                        End If

                    Next

                End If

            End If

            y1 += BarHeight
        Next

        'Draw Numbers
        Dim s As String
        Dim w, h, offset, inc As Short
        Dim LastWholeNumber As Integer

        If RulerIsHorizontal Then

            LastWholeNumber = (Math.Truncate(WidthInPixels / TickLocations(3)) * TickLocations(3))

            If NumberingIsLeftToRightHorizontal Then
                StartAt = 0
                EndAt = LastWholeNumber
                inc = 0
            Else
                StartAt = LastWholeNumber + TickLocations(3)
                EndAt = 0
                inc = -TickLocations(3)
            End If

            For x As Short = StartAt To EndAt Step TickLocations(3) * Direction

                s = CStr(CInt((CStr(inc) * gRulerScalingFactorSetByUser) * 10) / 10)

                inc += TickLocations(3)

                w = g.MeasureString(s, NumbersFont).Width

                If NumberingIsLeftToRightHorizontal Then
                    offset = 0
                Else
                    offset = WidthInPixels - LastWholeNumber + w
                End If

                g.DrawString(s, NumbersFont, New SolidBrush(gCurrentSkin.Numbers), (x + offset) * MagnifierFactor - w - MagnifierFactor * 0.5, NumbersWritenAt)

            Next

        Else

            LastWholeNumber = (Math.Truncate(HeightInPixels / TickLocations(3)) * TickLocations(3))

            If NumberingIsSmallOnTopToHighOnBottomVertical Then
                ' skip writing the bottom number when boxes are in the way
                If (HeightInPixels - LastWholeNumber < 28 / MagnifierFactor) Then LastWholeNumber -= TickLocations(3)
            End If

            If NumberingIsSmallOnTopToHighOnBottomVertical Then
                StartAt = TickLocations(3)
                EndAt = LastWholeNumber
                inc = TickLocations(3)
            Else
                StartAt = LastWholeNumber - TickLocations(3)
                EndAt = 0
                inc = TickLocations(3)
            End If

            For x As Short = StartAt To EndAt Step TickLocations(3) * Direction

                s = CStr(CInt((CStr(inc) * gRulerScalingFactorSetByUser) * 10) / 10)

                inc += TickLocations(3)

                w = g.MeasureString(s, NumbersFont).Width
                h = g.MeasureString(s, NumbersFont).Height / 2

                If NumberingIsSmallOnTopToHighOnBottomVertical Then
                    offset = 0
                Else
                    offset = HeightInPixels - LastWholeNumber + w
                End If

                If NumberingIsSmallOnTopToHighOnBottomVertical Then

                    If RulerHasScaleOnTopOrLeft Then
                        g.DrawString(s, NumbersFont, New SolidBrush(gCurrentSkin.Numbers), WidthInPixels - w, x * MagnifierFactor - h)
                    Else
                        g.DrawString(s, NumbersFont, New SolidBrush(gCurrentSkin.Numbers), 0, x * MagnifierFactor - h)
                    End If

                Else

                    If (x = EndAt) And ((HeightInPixels - LastWholeNumber) < (40 / MagnifierFactor)) Then
                        ' skip writing the top number when boxes are in the way
                    Else
                        If RulerHasScaleOnTopOrLeft Then
                            g.DrawString(s, NumbersFont, New SolidBrush(gCurrentSkin.Numbers), WidthInPixels - w, (HeightInPixels - inc + 100 - h) * MagnifierFactor)
                        Else
                            g.DrawString(s, NumbersFont, New SolidBrush(gCurrentSkin.Numbers), 0, (HeightInPixels - inc + 100 - h) * MagnifierFactor)
                        End If
                    End If

                End If

            Next

        End If

        If gShowLength Then ShowLengthOfRuler()

    End Sub

    'Private Sub ShowDividingLine()  ' testing here

    '    If MidPointIsOn Then
    '        Dim MidHeight As Integer = HeightInPixels / 2
    '        g.DrawLine(New Pen(gCurrentSkin.LineMidpoint), 0, MidHeight, WidthInPixels * MagnifierFactor, MidHeight)
    '    End If

    'End Sub

    Private Sub ShowLengthOfRuler()

        Dim h, w As Integer
        Dim s As String

        'draw green length number

        Try

            If RulerIsHorizontal Then

                s = CStr(CInt((CStr(WidthInPixels) * gRulerScalingFactorSetByUser) * 10) / 10)

                If NumberingIsLeftToRightHorizontal Then
                    w = g.MeasureString(s, SmallNumbersFont).Width

                    If RulerHasScaleOnTopOrLeft Then
                        g.DrawString(s, SmallNumbersFont, New SolidBrush(gCurrentSkin.LineLength), WidthInPixels * MagnifierFactor - w, 67)
                    Else
                        g.DrawString(s, SmallNumbersFont, New SolidBrush(gCurrentSkin.LineLength), WidthInPixels * MagnifierFactor - w, 21)
                    End If

                Else
                    If RulerHasScaleOnTopOrLeft Then
                        g.DrawString(s, SmallNumbersFont, New SolidBrush(gCurrentSkin.LineLength), 0, 59)
                    Else
                        g.DrawString(s, SmallNumbersFont, New SolidBrush(gCurrentSkin.LineLength), 0, 30)
                    End If
                End If

            Else

                s = CStr(CInt((CStr(HeightInPixels) * gRulerScalingFactorSetByUser) * 10) / 10)

                If NumberingIsSmallOnTopToHighOnBottomVertical Then

                    h = g.MeasureString(s, SmallNumbersFont).Height
                    If RulerHasScaleOnTopOrLeft Then
                        w = 81 - (g.MeasureString(s, SmallNumbersFont).Width)
                        g.DrawString(s, SmallNumbersFont, New SolidBrush(gCurrentSkin.LineLength), w, (HeightInPixels * MagnifierFactor - h))
                    Else
                        w = 19
                        g.DrawString(s, SmallNumbersFont, New SolidBrush(gCurrentSkin.LineLength), w, (HeightInPixels * MagnifierFactor - h))
                    End If

                Else
                    If RulerHasScaleOnTopOrLeft Then
                        w = 81 - (g.MeasureString(s, SmallNumbersFont).Width)
                        g.DrawString(s, SmallNumbersFont, New SolidBrush(gCurrentSkin.LineLength), w, 5)
                    Else
                        w = 19
                        g.DrawString(s, SmallNumbersFont, New SolidBrush(gCurrentSkin.LineLength), w, 5)
                    End If
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub DrawMeasuringLines()

        'Determine what lines are needed
        '  red line  = measuring line
        '  blue line = midpoint line
        '  golden line = golden ratio line

        Dim ARedLineIsNeeded As Boolean = True
        If RulerIsHorizontal Then
            If NumberingIsLeftToRightHorizontal Then
                If (MeasuringLineX < 0) OrElse (MeasuringLineX > WidthInPixels) Then ARedLineIsNeeded = False
            Else
                If (MeasuringLineXComplement < 0) OrElse (MeasuringLineXComplement > WidthInPixels) Then ARedLineIsNeeded = False
            End If
        Else
            If NumberingIsSmallOnTopToHighOnBottomVertical Then
                If (MeasuringLineY < 0) OrElse (MeasuringLineY > HeightInPixels) Then ARedLineIsNeeded = False
            Else
                If (MeasuringLineYComplement < 0) OrElse (MeasuringLineYComplement > HeightInPixels) Then ARedLineIsNeeded = False
            End If
        End If
        Dim ARedLineIsNotNeeded As Boolean = Not ARedLineIsNeeded

        Dim ABlueLineIsNeeded As Boolean = MidPointIsOn
        Dim ABlueLineIsNotNeeded As Boolean = Not ABlueLineIsNeeded

        Dim AGoldenLineIsNeeded As Boolean = GoldenRatioIsOn
        Dim AGoldenLineIsNotNeeded As Boolean = Not AGoldenLineIsNeeded

        Dim ThirdLinesAreNeeded As Boolean = ThirdsAreOn
        Dim ThirdLinesAreNotNeeded As Boolean = Not ThirdsAreOn

        If ARedLineIsNotNeeded AndAlso ABlueLineIsNotNeeded AndAlso AGoldenLineIsNotNeeded AndAlso ThirdLinesAreNotNeeded Then
            Exit Sub
        End If


        Dim r, b, g, t1, t2, RedLineSection, BlueLineSection, GoldenLineSection, T1Section, T2Section As Integer

        'determine red line section
        If ARedLineIsNeeded Then

            If RulerIsHorizontal Then

                If NumberingIsLeftToRightHorizontal Then
                    r = MeasuringLineX
                Else
                    r = MeasuringLineXComplement
                End If

                RedLineSection = Math.Truncate(r / 100)
                If ((WidthInPixels - (RedLineSection * 100)) < 98) Then RedLineSection -= 1
            Else

                If NumberingIsSmallOnTopToHighOnBottomVertical Then
                    r = MeasuringLineY
                Else
                    r = MeasuringLineYComplement
                End If

                RedLineSection = Math.Truncate(r / 100)
                If ((HeightInPixels - (RedLineSection * 100)) < 98) Then RedLineSection -= 1

            End If

        End If

        'determine blue line section
        If ABlueLineIsNeeded Then
            CalculateMidPoint()
            If RulerIsHorizontal Then
                b = MidpointLineX
            Else
                b = MidpointLineY
            End If
            BlueLineSection = Math.Truncate(b / 100)
        End If

        'determine golden line section
        If AGoldenLineIsNeeded Then

            g = CalculateGoldenLine()

            If (g < 121) AndAlso ((NumberingIsLeftToRightHorizontal AndAlso RulerIsHorizontal) OrElse (NumberingIsSmallOnTopToHighOnBottomVertical AndAlso RulerIsVertical)) Then
                GoldenLineSection = 0
            Else
                GoldenLineSection = Math.Truncate(g / 100)
            End If

        End If

        If ThirdLinesAreNeeded Then

            CalculateThirds(t1, t2)

            If (t1 < 121) AndAlso (RulerIsHorizontal OrElse (NumberingIsSmallOnTopToHighOnBottomVertical AndAlso RulerIsVertical)) Then
                T1Section = 0
            Else
                T1Section = Math.Truncate(t1 / 100)
            End If

            If (t2 < 121) AndAlso RulerIsHorizontal Then
                T2Section = 0

            ElseIf (t2 < 132) AndAlso RulerIsVertical Then
                T2Section = 0

            Else

                T2Section = Math.Truncate(t2 / 100)

            End If

        End If

        If RulerIsHorizontal AndAlso (Not NumberingIsLeftToRightHorizontal) Then
            If WidthInPixels <= 200 Then
                RedLineSection = 0
                BlueLineSection = 0
                T1Section = 0
                T2Section = 0
                GoldenLineSection = 0
            End If
        End If

        If RulerIsVertical AndAlso (Not NumberingIsSmallOnTopToHighOnBottomVertical) Then
            If HeightInPixels <= 200 Then
                RedLineSection = 0
                BlueLineSection = 0
                T1Section = 0
                T2Section = 0
                GoldenLineSection = 0
            End If
        End If

        '*****************************************************************************************************************************************************************

        'Determine overlapping lines
        Dim RedAndGoldOverlap As Boolean = (r = g)
        Dim RedAndBlueOverlap As Boolean = (r = b)
        Dim RedAndT1Overlap As Boolean = (r = t1)
        Dim RedAndT2Overlap As Boolean = (r = t2)

        '********************************************************

        If (r <> b) Then DrawALine(b, New SolidBrush(gCurrentSkin.LineMidpoint), New Pen(gCurrentSkin.LineMidpoint))
        If (r <> g) Then DrawALine(g, New SolidBrush(gCurrentSkin.LineGoldenRatio), New Pen(gCurrentSkin.LineGoldenRatio))
        If (r <> t1) Then DrawALine(t1, New SolidBrush(gCurrentSkin.LineThirds), New Pen(gCurrentSkin.LineThirds))
        If (r <> t2) Then DrawALine(t2, New SolidBrush(gCurrentSkin.LineThirds), New Pen(gCurrentSkin.LineThirds))
        If (r = b) OrElse (r = g) OrElse (r = t1) OrElse (r = t2) Then
            DrawALine(r, New SolidBrush(Color.Purple), New Pen(Color.Purple))
        Else
            DrawALine(r, New SolidBrush(gCurrentSkin.LineMeasuring), New Pen(gCurrentSkin.LineMeasuring))
        End If

        Dim EndPoint As Integer
        If RulerIsHorizontal Then
            EndPoint = WidthInPixels
        Else
            EndPoint = HeightInPixels
        End If

        Dim wColour As System.Drawing.Color
        Dim LinesInThisSection, w As Integer

        For Section As Integer = 0 To Math.Truncate(EndPoint / 100) + 1

            ResetQue()
            LinesInThisSection = 0

            'for every line that needs to be written, add its position on the ruler and its colour to the queue
            If ARedLineIsNeeded And (Section = RedLineSection) Then AddToQueue(r, gCurrentSkin.LineMeasuring) : LinesInThisSection += 1
            If ABlueLineIsNeeded And (Section = BlueLineSection) Then AddToQueue(b, gCurrentSkin.LineMidpoint) : LinesInThisSection += 1
            If AGoldenLineIsNeeded And (Section = GoldenLineSection) Then AddToQueue(g, gCurrentSkin.LineGoldenRatio) : LinesInThisSection += 1
            If ThirdLinesAreNeeded And (Section = T1Section) Then AddToQueue(t1, gCurrentSkin.LineThirds) : LinesInThisSection += 1
            If ThirdLinesAreNeeded And (Section = T2Section) Then AddToQueue(t2, gCurrentSkin.LineThirds) : LinesInThisSection += 1

            'remove all queued lines in the order
            For x As Integer = 1 To LinesInThisSection
                If (RulerIsHorizontal AndAlso NumberingIsLeftToRightHorizontal) OrElse (RulerIsVertical AndAlso NumberingIsSmallOnTopToHighOnBottomVertical) Then
                    PopLowestValueInQue(w, wColour)
                Else
                    PopHighestValueFromQueue(w, wColour)
                End If
                WritePixels(w, Section, LinesInThisSection, x, New SolidBrush(wColour))
            Next

        Next

    End Sub

    Friend Structure QueueStructure
        Friend Line As Integer
        Friend Colour As System.Drawing.Color
    End Structure
    Private MyQueue(6) As QueueStructure
    Private QueuePointer As Integer


    Private Sub ResetQue()
        QueuePointer = 0
        Array.Clear(MyQueue, 0, MyQueue.Length)
    End Sub

    Private Sub AddToQueue(ByVal Line As Integer, ByVal Colour As System.Drawing.Color)

        QueuePointer += 1
        MyQueue(QueuePointer).Line = Line
        MyQueue(QueuePointer).Colour = Colour

    End Sub

    Private Sub PopHighestValueFromQueue(ByRef PopLine As Integer, ByRef PopColour As System.Drawing.Color)

        'find the max value
        Dim max As Integer = Integer.MinValue
        For x As Integer = 0 To QueuePointer
            If MyQueue(x).Line > max Then max = MyQueue(x).Line
        Next

        'pop the max value
        For x As Integer = 1 To QueuePointer
            If MyQueue(x).Line = max Then
                PopLine = MyQueue(x).Line
                PopColour = MyQueue(x).Colour
                MyQueue(x).Line = -1
                Exit For
            End If
        Next

        RebuildQueue()
        QueuePointer -= 1

    End Sub

    Private Sub PopLowestValueInQue(ByRef PopLine As Integer, ByRef PopColour As System.Drawing.Color)

        'find the min value

        Dim min As Integer = Integer.MaxValue
        For x As Integer = 1 To QueuePointer
            If MyQueue(x).Line < min Then min = MyQueue(x).Line
        Next

        'pop the min value
        For x As Integer = 1 To QueuePointer
            If MyQueue(x).Line = min Then
                PopLine = MyQueue(x).Line
                PopColour = MyQueue(x).Colour
                MyQueue(x).Line = -1
                Exit For
            End If
        Next

        RebuildQueue()
        QueuePointer -= 1

    End Sub

    Private Sub RebuildQueue()

        'removes the -1 entry from the queue
        Dim y As Integer = 1
        For x As Integer = 1 To QueuePointer
            If MyQueue(x).Line = -1 Then
            Else
                MyQueue(y).Line = MyQueue(x).Line
                MyQueue(y).Colour = MyQueue(x).Colour
                y += 1
            End If
        Next

    End Sub

    Private Sub CalculateThirds(ByRef t1 As Integer, ByRef t2 As Integer)

        If RulerIsHorizontal Then
            t1 = WidthInPixels / 3
            t2 = WidthInPixels - t1
        Else
            t1 = HeightInPixels / 3
            t2 = HeightInPixels - t1
        End If

    End Sub

    Private Function CalculateGoldenLine() As Integer

        Const GoldenRatio As Double = 1.6180339887

        Dim GoldenRatioLength As Double
        If RulerIsHorizontal Then
            GoldenRatioLength = WidthInPixels / GoldenRatio
        Else
            GoldenRatioLength = HeightInPixels / GoldenRatio
        End If

        Dim iGoldenRatioLength As Integer = CType(GoldenRatioLength, Integer)

        Return iGoldenRatioLength

    End Function

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        'ensure on top of everything
        If gSuspendTimer2 Then
        Else
            MakeTopMost(Me.Handle.ToInt64)
        End If

    End Sub

    Private Sub DrawALine(ByVal WorkingDrawLineAt As Integer, ByVal WorkingBrushColour As System.Drawing.Brush, ByVal WorkingPenColour As System.Drawing.Pen)

        'Console.WriteLine("WorkingDrawLineAt {0} :", WorkingDrawLineAt)

        Try

            If RulerIsHorizontal Then

                If NumberingIsLeftToRightHorizontal Then

                    'Ruler is Horizontal and Numbering is Left to Right
                    If (WorkingDrawLineAt < 0) Or (WorkingDrawLineAt > WidthInPixels) Then Exit Sub
                    WorkingDrawLineAt -= 1

                    If RulerHasScaleOnTopOrLeft Then

                        'Ruler Is Horizontal and Numbering is Left to Right and Scale on Top
                        For x As Integer = 0 To MagnifierFactor - 1
                            g.DrawLine(WorkingPenColour, WorkingDrawLineAt * MagnifierFactor + x, 0, WorkingDrawLineAt * MagnifierFactor + x, 60)
                        Next x

                    Else

                        'Ruler Is Horizontal and Numbering is Left to Right and Scale on Bottom
                        For x As Integer = 0 To MagnifierFactor - 1
                            g.DrawLine(WorkingPenColour, WorkingDrawLineAt * MagnifierFactor + x, 40, WorkingDrawLineAt * MagnifierFactor + x, CType(HeightInPixels, Integer))
                        Next x

                    End If

                Else

                    'Ruler is Horizontal and Numbering is Reversed 
                    Dim WorkingLineXComplement As Integer = WorkingDrawLineAt
                    If (WorkingLineXComplement <= 0) Or (WorkingLineXComplement > WidthInPixels) Then Exit Sub

                    If RulerHasScaleOnTopOrLeft Then

                        'Ruler is Horizontal and Numbering is Reversed and Scale on Top

                        For x As Integer = 0 To MagnifierFactor - 1
                            'g.DrawLine(WorkingPenColour, (WidthInPixels - WorkingLineXComplement) * MagnifierFactor + x, 0, (WidthInPixels - WorkingLineXComplement) * MagnifierFactor + x, 60)
                            g.DrawLine(WorkingPenColour, WidthInPixels * MagnifierFactor - (WorkingLineXComplement * MagnifierFactor) + x, 0, (WidthInPixels - WorkingLineXComplement) * MagnifierFactor + x, 60)
                        Next x

                    Else

                        'Ruler is Horizontal and Numbering is Reversed and Scale on Bottom

                        For x As Integer = 0 To MagnifierFactor - 1
                            g.DrawLine(WorkingPenColour, (WidthInPixels - WorkingLineXComplement) * MagnifierFactor + x, 40, (WidthInPixels - WorkingLineXComplement) * MagnifierFactor + x, CType(HeightInPixels, Integer))
                        Next x

                    End If

                End If

            Else

                If NumberingIsSmallOnTopToHighOnBottomVertical Then

                    'Ruler is Vertical and Numbering is Left to Right

                    If (WorkingDrawLineAt < 0) Or (WorkingDrawLineAt > HeightInPixels) Then Exit Sub
                    WorkingDrawLineAt -= 1

                    If RulerHasScaleOnTopOrLeft Then

                        'Ruler Is Vertical and Numbering is Left to Right and Scale on Top

                        For y As Integer = 0 To MagnifierFactor - 1
                            g.DrawLine(WorkingPenColour, 0, WorkingDrawLineAt * MagnifierFactor + y, 60, WorkingDrawLineAt * MagnifierFactor + y)
                        Next y

                    Else

                        'Ruler Is Vertical and Numbering is Left to Right and Scale on Bottom

                        For y As Integer = 0 To MagnifierFactor - 1
                            g.DrawLine(WorkingPenColour, WidthInPixels - 60, WorkingDrawLineAt * MagnifierFactor + y, WidthInPixels, WorkingDrawLineAt * MagnifierFactor + y)
                        Next y

                    End If

                Else

                    'Ruler Is Vertical and Numbering is Reversed

                    Dim WorkingLineYComplement As Integer = WorkingDrawLineAt
                    If (WorkingLineYComplement <= 0) Or (WorkingLineYComplement > HeightInPixels) Then Exit Sub

                    'Console.WriteLine(WorkingLineYComplement)

                    If RulerHasScaleOnTopOrLeft Then

                        'Ruler Is Vertical and Numbering is Reversed and Scale on Top

                        For y As Integer = 0 To MagnifierFactor - 1
                            g.DrawLine(WorkingPenColour, 0, (HeightInPixels - WorkingLineYComplement) * MagnifierFactor + y, 60, (HeightInPixels - WorkingLineYComplement) * MagnifierFactor + y)
                        Next y


                    Else

                        For y As Integer = 0 To MagnifierFactor - 1
                            g.DrawLine(WorkingPenColour, WidthInPixels - 60, (HeightInPixels - WorkingLineYComplement) * MagnifierFactor + y, WidthInPixels, (HeightInPixels - WorkingLineYComplement) * MagnifierFactor + y)
                        Next y

                    End If

                End If

            End If

        Catch ex As Exception
            MsgBox(ex.TargetSite.Name & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error") : Me.DialogResult = 0
        End Try

    End Sub

    Private Sub WritePixels(ByVal PixelValue As Integer, ByVal Section As Integer, ByVal NumberOfEntriesInThisSection As Integer, ByVal EntryNumber As Integer, ByVal WorkingBrushColour As System.Drawing.Brush)

        Try

            If RulerIsHorizontal Then
                If (PixelValue <= 0) OrElse (PixelValue > WidthInPixels) Then Exit Sub
            Else
                If (PixelValue < 0) OrElse (PixelValue > HeightInPixels) Then Exit Sub
            End If

            Dim strPixelValue As String = String.Empty

            Dim ho, vo, w As Integer
            Dim LeftMargin As Integer
            Dim HeightForASingleLine As Integer
            Dim HeightForTheFirstOfTwoLines, HeightForTheSecondOfTwoLines As Integer
            Dim HeightForTheFirstOfThreeLines, HeightForTheSecondOfThreeLines, HeightForThirdOfThreeLines As Integer
            Dim HeightForFirstOfFourLines, HeightForSecondOfFourLines, HeightForThirdOfFourLines, HeightForForthOfFourLines As Integer

            Dim FontToUse As System.Drawing.Font = NumbersFont

            Dim Sectionx100xMagnifierFactor As Integer = Section * 100 * MagnifierFactor

            If RulerIsHorizontal Then

                If RulerHasScaleOnTopOrLeft Then

                    HeightForASingleLine = 67

                    HeightForTheFirstOfTwoLines = 58
                    HeightForTheSecondOfTwoLines = 77

                    HeightForTheFirstOfThreeLines = 60
                    HeightForTheSecondOfThreeLines = 72
                    HeightForThirdOfThreeLines = 84

                Else
                    HeightForASingleLine = 10

                    HeightForTheFirstOfTwoLines = 1
                    HeightForTheSecondOfTwoLines = 20

                    HeightForTheFirstOfThreeLines = 3
                    HeightForTheSecondOfThreeLines = 15
                    HeightForThirdOfThreeLines = 27

                End If

            Else

                If NumberingIsSmallOnTopToHighOnBottomVertical Then

                    HeightForASingleLine = Sectionx100xMagnifierFactor + 30 * MagnifierFactor

                    HeightForTheFirstOfTwoLines = Sectionx100xMagnifierFactor + 10 * MagnifierFactor
                    HeightForTheSecondOfTwoLines = HeightForTheFirstOfTwoLines + (40 * MagnifierFactor)

                    HeightForTheFirstOfThreeLines = Sectionx100xMagnifierFactor + 15 * MagnifierFactor
                    HeightForTheSecondOfThreeLines = HeightForTheFirstOfThreeLines + (25 * MagnifierFactor)
                    HeightForThirdOfThreeLines = HeightForTheSecondOfThreeLines + (25 * MagnifierFactor)

                    HeightForFirstOfFourLines = Sectionx100xMagnifierFactor + 30 * MagnifierFactor
                    HeightForSecondOfFourLines = Sectionx100xMagnifierFactor + 42 * MagnifierFactor
                    HeightForThirdOfFourLines = Sectionx100xMagnifierFactor + 54 * MagnifierFactor
                    HeightForForthOfFourLines = Sectionx100xMagnifierFactor + 66 * MagnifierFactor

                Else

                    HeightForASingleLine = (HeightInPixels * MagnifierFactor) - ((Section + 1) * 100 * MagnifierFactor) + (30 * MagnifierFactor)

                    HeightForTheFirstOfTwoLines = (HeightInPixels * MagnifierFactor) - ((Section + 1) * 100 * MagnifierFactor) + (10 * MagnifierFactor)
                    HeightForTheSecondOfTwoLines = HeightForTheFirstOfTwoLines + (40 * MagnifierFactor)

                    HeightForTheFirstOfThreeLines = (HeightInPixels * MagnifierFactor) - ((Section + 1) * 100 * MagnifierFactor) + (15 * MagnifierFactor)
                    HeightForTheSecondOfThreeLines = HeightForTheFirstOfThreeLines + (25 * MagnifierFactor)
                    HeightForThirdOfThreeLines = HeightForTheSecondOfThreeLines + (25 * MagnifierFactor)

                    HeightForFirstOfFourLines = Sectionx100xMagnifierFactor + 30 * MagnifierFactor
                    HeightForSecondOfFourLines = Sectionx100xMagnifierFactor + 42 * MagnifierFactor
                    HeightForThirdOfFourLines = Sectionx100xMagnifierFactor + 54 * MagnifierFactor
                    HeightForForthOfFourLines = Sectionx100xMagnifierFactor + 66 * MagnifierFactor

                End If

            End If


            Select Case NumberOfEntriesInThisSection

                '******************************************************************************************************************************

                Case Is = 1

                    Dim Offset As Integer = 0

                    If RulerIsHorizontal Then

                        If Section = 1 Then

                            If (WidthInPixels > 100) AndAlso (WidthInPixels <= 187) Then
                                If NumberingIsLeftToRightHorizontal Then
                                    Offset = -10
                                Else
                                    Offset = +40
                                End If
                            End If

                        End If

                    End If

                    If RulerIsVertical Then

                        If RulerHasScaleOnTopOrLeft Then
                            LeftMargin = 81
                        Else
                            LeftMargin = 19
                        End If

                        If PixelValue >= 1000 Then
                            FontToUse = SmallNumbersFont
                        End If

                    End If

                    strPixelValue = CStr(CInt((CStr(PixelValue) * gRulerScalingFactorSetByUser) * 10) / 10)

                    w = g.MeasureString(strPixelValue, FontToUse).Width
                    vo = g.MeasureString(strPixelValue, FontToUse).Height

                    ' The Number
                    If RulerIsHorizontal Then

                        If NumberingIsLeftToRightHorizontal Then
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, (Sectionx100xMagnifierFactor + 100 * MagnifierFactor / 2 - w / 2) + Offset * MagnifierFactor, HeightForASingleLine)
                        Else
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, (WidthInPixels * MagnifierFactor) - Sectionx100xMagnifierFactor - (100 * MagnifierFactor / 2 + w / 2) + Offset * MagnifierFactor, HeightForASingleLine)
                        End If

                    Else

                        vo = (vo * MagnifierFactor / 4)
                        If NumberingIsSmallOnTopToHighOnBottomVertical Then
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, LeftMargin - w / 2, HeightForASingleLine + vo)
                        Else
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, LeftMargin - w / 2, HeightForASingleLine + vo)
                        End If

                    End If

                    '******************************************************************************************************************************

                Case Is = 2

                    Dim Offset As Integer = 0

                    If RulerIsHorizontal Then

                        If (Section = 0) OrElse (WidthInPixels <= 220) Then
                            FontToUse = SmallNumbersFont
                            Offset = 22 * MagnifierFactor
                        End If

                        strPixelValue = CStr(CInt((CStr(PixelValue) * gRulerScalingFactorSetByUser) * 10) / 10)

                        w = g.MeasureString(strPixelValue, FontToUse).Width

                        If NumberingIsLeftToRightHorizontal Then

                            If EntryNumber = 1 Then
                                w = Sectionx100xMagnifierFactor + Offset * MagnifierFactor
                            Else
                                w = (Section + 1) * 100 * MagnifierFactor - w * MagnifierFactor - Offset * MagnifierFactor
                            End If

                        Else

                            If EntryNumber = 1 Then
                                w = WidthInPixels * MagnifierFactor - (Section + 1) * 100 * MagnifierFactor + Offset * MagnifierFactor
                            Else
                                w = WidthInPixels * MagnifierFactor - Sectionx100xMagnifierFactor - w * MagnifierFactor - Offset * MagnifierFactor
                            End If

                        End If

                        g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, w, HeightForASingleLine)

                    Else

                        If RulerHasScaleOnTopOrLeft Then
                            LeftMargin = 81
                        Else
                            LeftMargin = 19
                        End If

                        Dim SpecialOffsetA As Integer = 0
                        Dim SpecialOffsetB As Integer = 0
                        If HeightInPixels <= 215 Then
                            FontToUse = SmallNumbersFont
                            SpecialOffsetA = 8 * MagnifierFactor
                            SpecialOffsetB = 2 * MagnifierFactor
                        End If

                        If PixelValue >= 1000 Then
                            FontToUse = SmallNumbersFont
                        End If

                        If Section = 1 Then
                            If Not NumberingIsSmallOnTopToHighOnBottomVertical Then
                                FontToUse = SmallNumbersFont
                            End If
                        End If

                        strPixelValue = CStr(CInt((CStr(PixelValue) * gRulerScalingFactorSetByUser) * 10) / 10)

                        w = g.MeasureString(strPixelValue, FontToUse).Width

                        If EntryNumber = 1 Then
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, LeftMargin - w / 2, HeightForASingleLine + SpecialOffsetA)
                        Else
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, LeftMargin - w / 2, HeightForASingleLine + SpecialOffsetB + 25 * MagnifierFactor)
                        End If

                    End If

                    '******************************************************************************************************************************

                Case Is = 3

                    FontToUse = SmallNumbersFont

                    If RulerIsVertical Then
                        If RulerHasScaleOnTopOrLeft Then
                            LeftMargin = 77
                        Else
                            LeftMargin = 23
                        End If
                    End If

                    If RulerIsHorizontal Then

                        strPixelValue = CStr(CInt((CStr(PixelValue) * gRulerScalingFactorSetByUser) * 10) / 10)

                        w = g.MeasureString(strPixelValue, FontToUse).Width

                        If EntryNumber = 1 Then
                            vo = HeightForTheFirstOfThreeLines

                        ElseIf EntryNumber = 2 Then
                            vo = HeightForTheSecondOfThreeLines

                        Else
                            vo = HeightForThirdOfThreeLines

                        End If

                        If NumberingIsLeftToRightHorizontal Then
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, (Sectionx100xMagnifierFactor + 100 * MagnifierFactor / 2 - w / 2), vo)
                        Else
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, (WidthInPixels * MagnifierFactor) - ((Section + 1) * 100 * MagnifierFactor) + (100 * MagnifierFactor / 2 - w / 2), vo)
                        End If

                    Else


                        strPixelValue = CStr(CInt((CStr(PixelValue) * gRulerScalingFactorSetByUser) * 10) / 10)

                        If RulerIsVertical AndAlso (HeightInPixels <= 215) Then
                            'special case to make all numbers them fit
                            HeightForTheFirstOfThreeLines += 18
                            HeightForTheSecondOfThreeLines += 8
                        End If

                        If EntryNumber = 1 Then

                            w = g.MeasureString(strPixelValue, FontToUse).Width
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, LeftMargin - w / 2, HeightForTheFirstOfThreeLines)

                        ElseIf EntryNumber = 2 Then

                            w = g.MeasureString(strPixelValue, FontToUse).Width
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, LeftMargin - w / 2, HeightForTheSecondOfThreeLines)

                        Else

                            w = g.MeasureString(strPixelValue, FontToUse).Width
                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, LeftMargin - w / 2, HeightForThirdOfThreeLines)

                        End If

                    End If

                    '******************************************************************************************************************************

                Case Is = 4

                    If RulerIsHorizontal Then

                        If (WidthInPixels <= 217) Then
                            FontToUse = SmallNumbersFont
                            HeightForTheFirstOfTwoLines += 5
                            HeightForTheSecondOfTwoLines += 5
                        End If

                        strPixelValue = CStr(CInt((CStr(PixelValue) * gRulerScalingFactorSetByUser) * 10) / 10)

                        w = g.MeasureString(strPixelValue, FontToUse).Width

                        Dim HorizontalOffset As Integer = 0

                        If NumberingIsLeftToRightHorizontal Then
                            Select Case EntryNumber
                                Case Is = 1
                                    vo = HeightForTheFirstOfTwoLines
                                    If WidthInPixels > 215 Then
                                        HorizontalOffset = -25 * MagnifierFactor
                                    Else
                                        HorizontalOffset = -20 * MagnifierFactor
                                    End If
                                Case Is = 2
                                    vo = HeightForTheSecondOfTwoLines
                                    If WidthInPixels > 215 Then
                                        HorizontalOffset = -25 * MagnifierFactor
                                    Else
                                        HorizontalOffset = -20 * MagnifierFactor
                                    End If
                                Case Is = 3
                                    vo = HeightForTheFirstOfTwoLines
                                    If WidthInPixels > 215 Then
                                        HorizontalOffset = 25 * MagnifierFactor
                                    Else
                                        HorizontalOffset = 20 * MagnifierFactor
                                    End If
                                Case Is = 4
                                    vo = HeightForTheSecondOfTwoLines
                                    If WidthInPixels > 215 Then
                                        HorizontalOffset = 25 * MagnifierFactor
                                    Else
                                        HorizontalOffset = 20 * MagnifierFactor
                                    End If
                            End Select

                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, (Sectionx100xMagnifierFactor + 100 * MagnifierFactor / 2 - w / 2) + HorizontalOffset, vo)


                        Else

                            Select Case EntryNumber
                                Case Is = 1
                                    vo = HeightForTheFirstOfTwoLines
                                    If WidthInPixels > 217 Then
                                        HorizontalOffset = 0
                                    Else
                                        FontToUse = SmallNumbersFont
                                        HorizontalOffset = 20 * MagnifierFactor
                                    End If
                                Case Is = 2
                                    vo = HeightForTheSecondOfTwoLines
                                    If WidthInPixels > 217 Then
                                        HorizontalOffset = 0
                                    Else
                                        HorizontalOffset = 20 * MagnifierFactor
                                        FontToUse = SmallNumbersFont
                                    End If
                                Case Is = 3
                                    vo = HeightForTheFirstOfTwoLines
                                    If WidthInPixels > 217 Then
                                        HorizontalOffset = 50 * MagnifierFactor
                                    Else
                                        HorizontalOffset = 65 * MagnifierFactor
                                        FontToUse = SmallNumbersFont
                                    End If
                                Case Is = 4
                                    vo = HeightForTheSecondOfTwoLines
                                    If WidthInPixels > 217 Then
                                        HorizontalOffset = 50 * MagnifierFactor
                                    Else
                                        HorizontalOffset = 65 * MagnifierFactor
                                        FontToUse = SmallNumbersFont
                                    End If
                            End Select

                            Dim TotalSections = Math.Truncate(WidthInPixels / 100)
                            If (WidthInPixels Mod 100) > 0 Then TotalSections += 1

                            If (WidthInPixels <= 200) AndAlso (WidthInPixels > 100) Then
                                HorizontalOffset += 100
                            End If

                            g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, WidthInPixels * MagnifierFactor - ((TotalSections - Section) * 100 * MagnifierFactor) + HorizontalOffset, vo)

                        End If

                    Else

                        FontToUse = SmallNumbersFont

                        strPixelValue = CStr(CInt((CStr(PixelValue) * gRulerScalingFactorSetByUser) * 10) / 10)

                        w = g.MeasureString(strPixelValue, FontToUse).Width

                        Select Case EntryNumber
                            Case Is = 1
                                vo = HeightForFirstOfFourLines
                                ho = 10
                            Case Is = 2
                                vo = HeightForSecondOfFourLines
                                ho = 10
                            Case Is = 3
                                vo = HeightForThirdOfFourLines
                                ho = 10
                            Case Is = 4
                                vo = HeightForForthOfFourLines
                                ho = 10
                        End Select

                        If NumberingIsSmallOnTopToHighOnBottomVertical Then
                        Else
                            If HeightInPixels > 400 Then
                                vo = HeightInPixels * MagnifierFactor - 295 * MagnifierFactor + EntryNumber * 15
                            ElseIf HeightInPixels > 215 Then
                                vo = HeightInPixels * MagnifierFactor - 195 * MagnifierFactor + EntryNumber * 15
                            ElseIf HeightInPixels > 199 Then
                                vo = HeightInPixels * MagnifierFactor - 185 * MagnifierFactor + EntryNumber * 15
                            End If

                        End If

                        If RulerHasScaleOnTopOrLeft Then
                            ho += 60
                        End If

                        If Not NumberingIsSmallOnTopToHighOnBottomVertical Then
                            If HeightInPixels <= 200 Then
                                vo = HeightInPixels - 100 + vo
                            End If
                        End If

                        g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, ho, vo)

                    End If

                    '******************************************************************************************************************************

                Case Is = 5

                    FontToUse = SmallNumbersFont

                    strPixelValue = CStr(CInt((CStr(PixelValue) * gRulerScalingFactorSetByUser) * 10) / 10)

                    If RulerIsHorizontal Then

                        Select Case EntryNumber
                            Case Is = 1
                                ho = 20
                                vo = 6
                            Case Is = 2
                                ho = 20
                                vo = 25
                            Case Is = 3
                                ho = 42
                                vo = 17
                            Case Is = 4
                                ho = 65
                                vo = 6
                            Case Is = 5
                                ho = 65
                                vo = 25
                        End Select

                        If RulerHasScaleOnTopOrLeft Then vo += 57

                        If NumberingIsLeftToRightHorizontal Then
                            Select Case EntryNumber
                                Case Is = 1
                                    ho += 2
                                Case Is = 2
                                    ho += 2
                                Case Is = 4
                                    ho -= 3
                                Case Is = 5
                                    ho -= 3
                            End Select
                        End If
                        g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, ho * MagnifierFactor, vo)

                    Else

                        Select Case EntryNumber
                            Case Is = 1
                                ho = 0
                                vo = 35
                            Case Is = 4
                                ho = 0
                                vo = 60
                            Case Is = 3
                                ho = 12
                                vo = 47
                            Case Is = 2
                                ho = 25
                                vo = 35
                            Case Is = 5
                                ho = 25
                                vo = 60
                        End Select

                        If NumberingIsSmallOnTopToHighOnBottomVertical Then
                        Else

                            If HeightInPixels <= 200 Then
                                vo = HeightInPixels - 100 + vo
                            End If

                            If HeightInPixels = 299 Then
                                vo += 100
                                If RulerHasScaleOnTopOrLeft Then
                                    ho -= 6
                                End If
                            End If

                        End If

                        If RulerHasScaleOnTopOrLeft Then ho += 60
                        g.DrawString(strPixelValue, FontToUse, WorkingBrushColour, ho, vo * MagnifierFactor)

                    End If

            End Select

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub RulerPanel_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles Me.Scroll

        AutoScrollPositionX = sender.autoscrollposition.x * -1
        AutoScrollPositionY = sender.autoscrollposition.y * -1

    End Sub

    Private Sub RulerPanel_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RulerPanel.MouseLeave

        MoveFlag = False

        If gAutoExpand Then

            Dim ReBuildNeeded As Boolean = False

            If gCollapseCount > 0 Then

                For x As Integer = 0 To gCollapseCount
                    MakeReadingGuideThinner()
                Next
                gCollapseCount = 0

                ReBuildNeeded = True

            End If

            If gBoxScale = gBigBox Then

                gBoxScale = 10
                gOldBoxScale = gBoxScale

                ReBuildNeeded = True

            End If

            If ReBuildNeeded Then

                SetBoxSizes()
                RulerDrawBackground()
                RulerDrawAllLines()

            End If

            If gHoldShowLengthSet Then

                gHoldShowLengthSet = False
                gShowLength = gHoldShowLength

            End If

        End If

    End Sub



    Private Sub frmRuler_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RulerPanel.MouseMove

        Static Dim lastmessage As String
        Dim masterCommentForAllMouseButtons As String = String.Empty
        Dim leftMouseButtonComment As String = String.Empty
        Dim rightMouseButtonComment As String = String.Empty
        Dim middleMouseButtonComment As String = String.Empty

        Try

            'V3.4 added changes for auto expand buttons

            If gAutoExpand Then

                If CheckIn(e.Location, MasterBoxA) OrElse CheckIn(e.Location, MasterBoxB) Then

                    gBoxScale = gBigBox

                    If gHoldShowLengthSet Then
                    Else
                        gHoldShowLengthSet = True
                        gHoldShowLength = gShowLength
                        gShowLength = False
                    End If

                Else

                    If gCollapseCount = 0 Then
                        ' this ensures the box is only grown and shrunk once while the curson is on the ruler panel
                        gBoxScale = gNormalBox
                    End If

                    If gHoldShowLengthSet Then
                        gHoldShowLengthSet = False
                        gShowLength = gHoldShowLength
                    End If

                End If

            Else

                gBoxScale = gNormalBox

            End If


            If gBoxScale = gOldBoxScale Then

            Else

                If gShowReadingGuide Then

                    SetBoxSizes()
                    RulerDrawBackground()
                    RulerDrawAllLines()

                    If gOldBoxScale = gNormalBox Then

                        If gCollapseCount = 0 Then

                            While AdjustReadingGuideThickness < 30
                                MakeReadingGuideThicker()
                                gCollapseCount += 1
                            End While

                        End If

                    Else

                        For x As Integer = 0 To gCollapseCount
                            MakeReadingGuideThinner()
                        Next
                        gCollapseCount = 0

                    End If

                End If

                SetBoxSizes()
                RulerDrawBackground()
                RulerDrawAllLines()

                gOldBoxScale = gBoxScale

            End If

            ' v3.4 end changes 


            If gShowReadingGuide Then

                If CheckIn(e.Location, MasterBoxA) OrElse CheckIn(e.Location, MasterBoxB) Then

                    If ControlsAreHidden Then
                        ControlsAreHidden = False
                        Me.Refresh()

                        If Timer1.Enabled Then
                        Else
                            Timer1.Enabled = True
                            Timer1.Start()
                        End If

                    End If

                Else

                    If ControlsAreHidden Then
                    Else
                        If Timer1.Enabled Then
                        Else
                            Timer1.Enabled = True
                            Timer1.Start()
                        End If
                    End If

                End If

            End If

            If CheckIn(e.Location, MasterBoxA) Then

                If CheckIn(e.Location, NumberBox) Then

                    masterCommentForAllMouseButtons = gResourceManager.GetString("0003") 'Set ruler length

                ElseIf CheckIn(e.Location, ResizeLeftOrUpBox) Then

                    If RulerIsHorizontal Then
                        leftMouseButtonComment = gResourceManager.GetString("0004") 'Resize ruler to the left
                        middleMouseButtonComment = gResourceManager.GetString("0081") 'Auto-resize
                        rightMouseButtonComment = gResourceManager.GetString("0005") 'Lock left
                    Else
                        leftMouseButtonComment = gResourceManager.GetString("0006") 'Resize ruler upwards
                        middleMouseButtonComment = gResourceManager.GetString("0081") 'Auto-resize
                        rightMouseButtonComment = gResourceManager.GetString("0007") 'Lock top
                    End If

                ElseIf CheckIn(e.Location, TickPositionBox) Then

                    If gShowReadingGuide Then

                        masterCommentForAllMouseButtons = gResourceManager.GetString("0008") 'Switch from reading guide to ruler

                    Else

                        If RulerIsHorizontal Then
                            If RulerHasScaleOnTopOrLeft Then  ' this is where the ticks will be 
                                leftMouseButtonComment = gResourceManager.GetString("0009") 'Put ticks on bottom of the ruler
                            Else
                                leftMouseButtonComment = gResourceManager.GetString("0010") 'Put ticks on top of the ruler
                            End If

                        Else

                            If RulerHasScaleOnTopOrLeft Then  ' this is where the ticks will be 
                                leftMouseButtonComment = gResourceManager.GetString("0012") 'Put ticks on the left hand side of the ruler
                            Else
                                leftMouseButtonComment = gResourceManager.GetString("0011") 'Put ticks on the right hand side of the ruler
                            End If
                        End If

                        rightMouseButtonComment = gResourceManager.GetString("0013") 'Switch from ruler to reading guide

                    End If

                ElseIf CheckIn(e.Location, HorizontalVeriticalBox) Then
                    If RulerIsHorizontal Then
                        masterCommentForAllMouseButtons = gResourceManager.GetString("0014") 'Make ruler vertical
                    Else
                        masterCommentForAllMouseButtons = gResourceManager.GetString("0015") 'Make ruler horizontal
                    End If

                ElseIf CheckIn(e.Location, BackgroundBox) Then

                    Dim lCurrentSkinName As String = TranslateASkinNameFromEnglishtToAnotherLanguage(gCurrentSkin.Name)
                    Dim lNextEnabledSkinName As String = TranslateASkinNameFromEnglishtToAnotherLanguage(gNextEnabledSkinName)
                    Dim lPreviousEnabledSkinName As String = TranslateASkinNameFromEnglishtToAnotherLanguage(gPreviousEnabledSkinName)


                    If gCurrentSkin.Name = gNextEnabledSkinName Then
                        masterCommentForAllMouseButtons = gResourceManager.GetString("0016").Replace("{0}", lNextEnabledSkinName) '"The '" & gNextEnabledSkinName & "' skin is the only available enabled skin."
                    Else
                        If gNextEnabledSkinName = gPreviousEnabledSkinName Then
                            masterCommentForAllMouseButtons = gResourceManager.GetString("0017").Replace("{0}", lPreviousEnabledSkinName) '"Click for the '" & gPreviousEnabledSkinName & "' skin"
                        Else
                            masterCommentForAllMouseButtons = gResourceManager.GetString("0018").Replace("{0}", lCurrentSkinName) '"Current skin is '" & gCurrentSkin.Name & "'. Change to:"
                            leftMouseButtonComment = "'" & lPreviousEnabledSkinName & "'"
                            rightMouseButtonComment = "'" & lNextEnabledSkinName & "'"

                        End If

                    End If

                ElseIf CheckIn(e.Location, MagnifyBox) Then
                    masterCommentForAllMouseButtons = gResourceManager.GetString("0019")  '"Start/Stop Microsoft's Magnify program"

                End If

            ElseIf CheckIn(e.Location, MasterBoxB) Then

                If CheckIn(e.Location, ResizeRightOrDownBox) Then
                    If RulerIsHorizontal Then
                        leftMouseButtonComment = gResourceManager.GetString("0020") '"Resize ruler to the right"
                        middleMouseButtonComment = gResourceManager.GetString("0081") 'Auto-resize
                        rightMouseButtonComment = gResourceManager.GetString("0021") '"Lock right"
                    Else
                        leftMouseButtonComment = gResourceManager.GetString("0022") '"Resize ruler downwards"
                        middleMouseButtonComment = gResourceManager.GetString("0081") 'Auto-resize
                        rightMouseButtonComment = gResourceManager.GetString("0023") '"Lock bottom"
                    End If

                ElseIf CheckIn(e.Location, AboutBox) Then

                    masterCommentForAllMouseButtons = gResourceManager.GetString("0024") '"Options - Help - About"

                ElseIf CheckIn(e.Location, MinimizeNowBox) Then
                    masterCommentForAllMouseButtons = gResourceManager.GetString("0025") '"Minimize"

                ElseIf CheckIn(e.Location, CloseBox) Then
                    masterCommentForAllMouseButtons = gResourceManager.GetString("0026") '"Exit"

                End If

            Else

                masterCommentForAllMouseButtons = ""
                rightMouseButtonComment = ""
                leftMouseButtonComment = ""
                lastmessage = masterCommentForAllMouseButtons
                ToolTip1.SetToolTip(RulerPanel, masterCommentForAllMouseButtons)
                Exit Sub

            End If

            If gShowReadingGuide Then

                If masterCommentForAllMouseButtons = gResourceManager.GetString("0027") Then '"Switch from reading guide to ruler" 
                Else
                    masterCommentForAllMouseButtons = masterCommentForAllMouseButtons.Replace(gResourceManager.GetString("0028"), gResourceManager.GetString("0029")) 'Replace("ruler", "reading guide")
                    leftMouseButtonComment = leftMouseButtonComment.Replace(gResourceManager.GetString("0028"), gResourceManager.GetString("0029")) 'Replace("ruler", "reading guide")
                    rightMouseButtonComment = rightMouseButtonComment.Replace(gResourceManager.GetString("0028"), gResourceManager.GetString("0029")) 'Replace("ruler", "reading guide"), gResourceManager.GetString("0029")) 'Replace("ruler", "reading guide")
                End If

            End If

            If leftMouseButtonComment <> String.Empty Then leftMouseButtonComment &= " " & gResourceManager.GetString("0030") ' " (left click)"
            If middleMouseButtonComment <> String.Empty Then middleMouseButtonComment &= " " & gResourceManager.GetString("0082") ' " (middle click)"
            If rightMouseButtonComment <> String.Empty Then rightMouseButtonComment &= " " & gResourceManager.GetString("0031") ' " (right click)"

            masterCommentForAllMouseButtons &= vbCrLf
            leftMouseButtonComment &= vbCrLf
            If middleMouseButtonComment <> String.Empty Then
                middleMouseButtonComment &= vbCrLf
            End If
            rightMouseButtonComment &= vbCrLf

            masterCommentForAllMouseButtons = masterCommentForAllMouseButtons & leftMouseButtonComment & middleMouseButtonComment & rightMouseButtonComment
            masterCommentForAllMouseButtons = masterCommentForAllMouseButtons.Trim

            If (masterCommentForAllMouseButtons = lastmessage) Then
            Else
                ToolTip1.SetToolTip(RulerPanel, masterCommentForAllMouseButtons)
                lastmessage = masterCommentForAllMouseButtons
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Flush()

        Try

            Dim x As Int32 = Int(System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / 1024)
            If x < 5000 Then Exit Try

            If System.Environment.OSVersion.Platform = PlatformID.Win32NT Then
                Dim p As Process = Process.GetCurrentProcess
                Dim Dummy As Integer = SafeNativeMethods.SetProcessWorkingSetSize(p.Handle, -1, -1)
                p.Dispose()
            Else
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub frmRuler_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LostFocus

        DefaultPenColour = Pens.Gray
        RulerDrawAllLines()
        DefaultPenColour = Pens.Black

    End Sub

    Private Sub SetSettingsVersion()

        Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim appVersion As Version = a.GetName().Version
        Dim appVersionString As String = appVersion.ToString

        If My.Settings.ApplicationVersion <> appVersion.ToString Then
            UgradeToStronglyNamedSettings()
            My.Settings.Upgrade()
            My.Settings.ApplicationVersion = appVersionString
        End If

    End Sub


    Private Sub UgradeToStronglyNamedSettings()

        ' note this routine works on systems where the application has been deployed in either a release or a debug configuration on a machine, but not both

        Try

            Dim MyProgram As String = My.Application.Info.AssemblyName

            Dim StrongNameSettingsFolderForThisVersion As String = Path.GetDirectoryName(System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath)

            'if the Strong Name settings folder for this version already exists then we have already converted a Non Strong Name settings folder if one had existed and our job here is done
            If Directory.Exists(StrongNameSettingsFolderForThisVersion) Then Exit Sub

            Dim ParentOfStrongNameSettingsFolderForThisVersion As String = Path.GetDirectoryName(StrongNameSettingsFolderForThisVersion)

            'if the Parent of the Strong Name settings folder for this version already exists then we have already converted a Non Strong Name settings folder if one had existed and our job here is done
            If Directory.Exists(ParentOfStrongNameSettingsFolderForThisVersion) Then Exit Sub

            Dim GrandParentOfStrongNameSettingsFolderForThisVersion As String = StrongNameSettingsFolderForThisVersion.Remove(StrongNameSettingsFolderForThisVersion.IndexOf("\" & MyProgram & ".exe_StrongName_"))

            ' in the grand parent of the strong name folder for this version there should be either:
            ' no children with a non strong name for this program
            ' or just one child with a non strong name for this program 

            ' if there are no children we are done
            ' if there is a child we need to copy it over to the new strong name folder so that it can be upgraded

            Dim ParentOfNonStrongNameSettingsFolder As String = String.Empty

            Dim dir As New DirectoryInfo(GrandParentOfStrongNameSettingsFolderForThisVersion)

            For Each IndividualDirectory As DirectoryInfo In dir.GetDirectories()

                If IndividualDirectory.FullName.Contains(MyProgram & ".exe_Url_") Then

                    ParentOfNonStrongNameSettingsFolder = IndividualDirectory.FullName
                    Exit For

                End If

            Next

            ' if there were no children with non strong names we are done
            If ParentOfNonStrongNameSettingsFolder = String.Empty Then Exit Sub

            ' there is a child so we need to copy it over to the new strong name folder in order that may be upgraded

            ' Find the non strong name directory with the greatest version number

            Dim NonStrongNameSettingsFolderWithGreatestVersionNumber As String = String.Empty

            dir = New DirectoryInfo(ParentOfNonStrongNameSettingsFolder)

            For Each IndividualDirectory As DirectoryInfo In dir.GetDirectories()

                If IndividualDirectory.FullName > NonStrongNameSettingsFolderWithGreatestVersionNumber Then

                    NonStrongNameSettingsFolderWithGreatestVersionNumber = IndividualDirectory.FullName

                End If

            Next

            'get the version number of the non strong name directory with the greatest version number

            Dim VersionNumberOfNonStrongNameSettingsFolderWithGreatestVersionNumber As String = NonStrongNameSettingsFolderWithGreatestVersionNumber.Remove(0, NonStrongNameSettingsFolderWithGreatestVersionNumber.LastIndexOf("\") + 1)

            Dim NewStrongNameSettingsFolderWithGreatestVersionNumber = ParentOfStrongNameSettingsFolderForThisVersion & "\" & VersionNumberOfNonStrongNameSettingsFolderWithGreatestVersionNumber

            If Directory.Exists(NewStrongNameSettingsFolderWithGreatestVersionNumber) Then

                ' this should not happen

            Else

                '  copy over the old non strong name settings information to the new strong name folder so that it can be upgraded

                Directory.CreateDirectory(NewStrongNameSettingsFolderWithGreatestVersionNumber)
                CopyDirectory(NonStrongNameSettingsFolderWithGreatestVersionNumber, NewStrongNameSettingsFolderWithGreatestVersionNumber, True)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CopyDirectory(sourceDir As String, destinationDir As String, recursive As Boolean)

        'the code for this subroutine was extracted from:
        'https://docs.microsoft.com/en-us/dotnet/standard/io/how-to-copy-directories

        ' Get information about the source directory
        Dim dir As New DirectoryInfo(sourceDir)

        ' Check if the source directory exists
        If Not dir.Exists Then
            Throw New DirectoryNotFoundException($"Source directory not found: {dir.FullName}")
        End If

        ' Cache directories before we start copying
        Dim dirs As DirectoryInfo() = dir.GetDirectories()

        ' Create the destination directory
        Directory.CreateDirectory(destinationDir)

        ' Get the files in the source directory and copy to the destination directory
        For Each file As FileInfo In dir.GetFiles()
            Dim targetFilePath As String = Path.Combine(destinationDir, file.Name)
            file.CopyTo(targetFilePath)
        Next

        ' If recursive and copying subdirectories, recursively call this method
        If recursive Then
            For Each subDir As DirectoryInfo In dirs
                Dim newDestinationDir As String = Path.Combine(destinationDir, subDir.Name)
                CopyDirectory(subDir.FullName, newDestinationDir, True)
            Next
        End If

    End Sub



    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        ' auto hide controls

        Dim MousePositionOnRulerX As Integer = MousePosition.X - (RulerPanel.Location.X - gPrimaryScreenHome.X)
        Dim MousePositionOnRulerY As Integer = MousePosition.Y - (RulerPanel.Location.Y - gPrimaryScreenHome.Y)
        Dim RelativeMousePosition As New Point(MousePositionOnRulerX, MousePositionOnRulerY)

        If CheckIn(RelativeMousePosition, MasterBoxA) OrElse CheckIn(RelativeMousePosition, MasterBoxB) Then
            ControlsAreHidden = False
        Else

            Timer1.Stop()
            Timer1.Enabled = False
            ControlsAreHidden = True

            'added in support of a reading guide thinner than the minimum needed to display the boxes
            If lHoldThickness > 0 Then

                If gAutoExpand Then
                    If RulerIsHorizontal Then
                        AdjustReadingGuideThickness = lHoldThickness - 20
                    Else
                        AdjustReadingGuideThickness = lHoldThickness - 10
                    End If
                Else
                    AdjustReadingGuideThickness = lHoldThickness - 10
                End If

                lHoldThickness = 0

                If AdjustReadingGuideThickness < 0 Then
                    AdjustReadingGuideThickness = 0
                End If

                SetupForRulerOrReadingGuide()
                RulerDrawBackground()

            Else

                Me.Refresh()

            End If

        End If

    End Sub



#Region "Manage Instances"

    Private Sub CreateAMutex()

        Try

            Dim Dummy As Integer = SafeNativeMethods.CreateMutex(0&, 0&, "AR4WMutex")

        Catch ex As Exception
        End Try

    End Sub

#End Region

    'added in v3.4.2 ; handles sitution when ruler is closed from taskbar

    Friend CloseOnlyOnce As Boolean = True
    Friend CloseOnDemand As Boolean = False
    Private Sub frmRuler_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing


        If CloseOnlyOnce Then
            FrmRuler_Close()
        End If

    End Sub


End Class
