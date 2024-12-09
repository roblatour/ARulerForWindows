Imports System
Imports System.Runtime.InteropServices
Imports System.Security

Imports System.Globalization
Imports System.Reflection
Imports System.Resources

Imports System.IO
Imports System.Drawing.Imaging
Imports System.Text
Imports System.Diagnostics.Eventing.Reader
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Window

Module modCommon
    Friend Structure SkinStructure
        Friend Name As String
        Friend SkinEnabled As Boolean
        Friend LineLength As Color
        Friend LineMeasuring As Color
        Friend LineMidpoint As Color
        Friend LineThirds As Color
        Friend LineGoldenRatio As Color
        Friend LineBoxes As Color
        Friend LineTicksAndSides As Color
        Friend Numbers As Color
        Friend TransparencyColour As Color
        Friend Opacity As Single
        Friend Tiled As Boolean
        Friend HorizontalRotation As Integer
        Friend VerticalRotation As Integer
        Friend HorizontalBackground As Image
        Friend VerticalBackground As Image
        Friend BackgroundImageFileName As String
    End Structure

    Friend XML_Path_Name As String
    Friend Const MyExtention As String = ".ar4w"

    Friend gRulerScalingFactorSetByUser As Single = 1 '.25
    Friend gNoLines As Boolean = False

    Friend MyDefaultTransparentColour As Color = System.Drawing.Color.FromArgb(255, 250, 250, 250)
    Friend MyTransparentColour As Color = System.Drawing.Color.FromArgb(255, 250, 250, 250)

    Friend gKey_Clear As String = String.Empty
    Friend gKey_ShowLength As String = String.Empty
    Friend gKey_ShowMidPoint As String = String.Empty
    Friend gKey_ShowThirds As String = String.Empty
    Friend gKey_ShowGoldenRatioLine As String = String.Empty
    Friend gKey_ReverseNumbers As String = String.Empty
    Friend gKey_SwitchBetweenRulerAndGuide As String = String.Empty

    Friend gKey_PageUpPageDown As String = String.Empty
    Friend gKey_Home As String = String.Empty
    Friend gKey_Escape As String = String.Empty

    Friend gLicenseWebSite As String = String.Empty

    Friend gRefreshRulerAndAboutWindow As Boolean = False

    Friend gReloadHelp As Boolean = False
    Friend gPreferredLocationNotSet As Point = New Point(-1, -1)
    Friend gPeferedLocationForHelp As Point = New Point(-1, -1)

    Friend gCurrentSkin As SkinStructure
    Friend gAllSkinNames() As String
    Friend gNextEnabledSkinName As String = ""
    Friend gPreviousEnabledSkinName As String = ""

    Friend gMaxWidth As Integer = 0
    Friend gMaxHeight As Integer = 0

    Friend gMinX As Integer = 0
    Friend gMinY As Integer = 0

    Friend gPrimaryScreenHome As Point

    Friend frmWhiteBackGround As frmWhiteBackGround = New frmWhiteBackGround

    Friend MagnifierFactor As Integer = 1
    Friend NewSize As Single = 0
    Friend PanelLocationX As Integer = 0
    Friend PanelLocationY As Integer = 0
    Public ScreenCapture As Image = Nothing
    Friend RulerIsHorizontal As Boolean = True
    Friend ControlsAreHidden As Boolean = False
    Friend RulerIsVertical As Boolean = Not RulerIsHorizontal

    Friend gSameLocationAsSkinWindow As Point

    Friend Enum BackGroundChoices
        Wood = 1
        Steel = 2
        Plastic = 3
        Yellow = 4
    End Enum
    'Friend BackGroundChoice As BackGroundChoices = BackGroundChoices.Wood

    Friend gShowBalloons As Boolean = True
    Friend gQuickStartup As Boolean = False
    Friend gFenceRulerOnScreen As Boolean = True
    Friend gAutoExpand As Boolean = False
    Friend gShowLength As Boolean = True
    Friend gShowReadingGuide As Boolean = False
    Friend gBump As Integer = 15
    Friend gOKToExitNow As Boolean = False

    Friend gLockLeft As Boolean = False
    Friend gLockRight As Boolean = False
    Friend gLockTop As Boolean = False
    Friend gLockBottom As Boolean = False

    Friend gSetLocation As Point = New Point(0, 0)
    Friend gRulerPanelBounds As Rectangle

    Friend gUniversaleScale As Single = 1 'v 3.1

    Friend NumbersFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Friend SmallNumbersFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ' Friend WholeNumberFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    'Friend FractionsFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

    Friend DefaultPenColour As System.Drawing.Pen = Pens.Black
    Friend DefaultBrushColour As System.Drawing.Brush = Brushes.Black

    Friend gCurrentBaseBitMap As Bitmap

    Friend gTempDateFileName As String = System.Environment.GetEnvironmentVariable("TEMP") & "\~ccrl.tmp"

    Friend gScreens As Screen()

    Friend gResourceManager As ResourceManager

    Friend Const gBigBox As Integer = 20
    Friend Const gNormalBox As Integer = 10
    Friend gBoxScale = gNormalBox

    Friend gSuspendTimer2 As Boolean = False

    Friend gAboutWidth As Integer = 948
    Friend gAboutHeight As Integer = 540

    Friend Function ResizeMe(ByVal CurrentSize As Size) As Size  'v3.1

        If gUniversaleScale = 1 Then
            Return CurrentSize
        Else
            Return New Size(CurrentSize.Width * gUniversaleScale, CurrentSize.Height * gUniversaleScale)
        End If

    End Function

    Friend Sub MakeTopMost(ByVal hwnd As Integer, Optional ByVal MakeMeTopMostFlag As Boolean = True)

        Dim HWND_TOPMOST As Integer = -1
        Dim SWP_NOMOVE As Integer = &H2
        Dim SWP_NOSIZE As Integer = &H1
        Dim TOPMOST_FLAGS As Integer = SWP_NOMOVE Or SWP_NOSIZE

        If MakeMeTopMostFlag Then
            HWND_TOPMOST = -1
        Else
            HWND_TOPMOST = 0
        End If

        Try
            Dim Dummy As Integer = SafeNativeMethods.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
        Catch ex As Exception
        End Try

    End Sub

    Friend Sub StartAProcess(ByVal ProcessToRun As String)

        Try

            Dim myProcess As New Process
            myProcess.StartInfo.Verb = "open"
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.FileName = ProcessToRun
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.Start()
            Application.DoEvents()

        Catch ex As Exception
        End Try

    End Sub

    'Friend Sub DrawAnEmptyBox(ByVal g As Graphics, ByVal Box As Rectangle)
    '    g.DrawRectangle(DefaultPenColour, Box)
    'End Sub

    Friend Sub DrawAnXInABox(ByVal g As Graphics, ByVal Box As Rectangle)
        Dim BigBoxOffset As Integer = IIf(gBoxScale = gBigBox, 5, 0)
        g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), Box)
        g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + BigBoxOffset, Box.Y + BigBoxOffset, Box.X + 10 + BigBoxOffset, Box.Y + 10 + BigBoxOffset)
        g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 10 + BigBoxOffset, Box.Y + BigBoxOffset, Box.X + BigBoxOffset, Box.Y + 10 + BigBoxOffset)
    End Sub

    Friend Sub DrawAQuestionMarkInABox(ByVal g As Graphics, ByVal Box As Rectangle)

        Dim BigBoxOffset As Integer = IIf(gBoxScale = gBigBox, 5, 0)

        If gBoxScale = gBigBox Then

            Dim pts() As Point = {
            New Point(Box.X + 3 + BigBoxOffset, Box.Y + 3 + BigBoxOffset),
            New Point(Box.X + 3 + BigBoxOffset, Box.Y + 2 + BigBoxOffset),
            New Point(Box.X + 5 + BigBoxOffset, Box.Y + 1 + BigBoxOffset),
            New Point(Box.X + 7 + BigBoxOffset, Box.Y + 2 + BigBoxOffset),
            New Point(Box.X + 5 + BigBoxOffset, Box.Y + 5 + BigBoxOffset),
            New Point(Box.X + 5 + BigBoxOffset, Box.Y + 7 + BigBoxOffset)
            }

            g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), Box)
            g.DrawCurve(New Pen(gCurrentSkin.LineBoxes), pts)
            g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 5 + BigBoxOffset, Box.Y + 10 + BigBoxOffset, Box.X + 5 + BigBoxOffset, Box.Y + 11 + BigBoxOffset)

        Else

            Dim pts() As Point = {
            New Point(Box.X + 3 + BigBoxOffset, Box.Y + 4 + BigBoxOffset),
            New Point(Box.X + 3 + BigBoxOffset, Box.Y + 3 + BigBoxOffset),
            New Point(Box.X + 5 + BigBoxOffset, Box.Y + 2 + BigBoxOffset),
            New Point(Box.X + 7 + BigBoxOffset, Box.Y + 3 + BigBoxOffset),
            New Point(Box.X + 5 + BigBoxOffset, Box.Y + 6 + BigBoxOffset),
            New Point(Box.X + 5 + BigBoxOffset, Box.Y + 7 + BigBoxOffset)
            }

            g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), Box)
            g.DrawCurve(New Pen(gCurrentSkin.LineBoxes), pts)
            g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 5 + BigBoxOffset, Box.Y + 9 + BigBoxOffset, Box.X + 5 + BigBoxOffset, Box.Y + 10 + BigBoxOffset)

        End If

    End Sub

    Friend Sub DrawASquareInABox(ByVal g As Graphics, ByVal Box As Rectangle)

        Dim BigBoxOffset As Integer = IIf(gBoxScale = gBigBox, 5, 0)
        g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), Box)
        g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), New Rectangle(Box.X + 2 + BigBoxOffset, Box.Y + 2 + BigBoxOffset, 6, 6))
    End Sub

    Friend Sub DrawALockedBox(ByVal g As Graphics, ByVal Box As Rectangle)

        Dim BigBoxOffset As Integer = IIf(gBoxScale = gBigBox, 5, 0)
        g.FillRectangle(Brushes.Tomato, New Rectangle(Box.X + 3 + BigBoxOffset, Box.Y + 3 + BigBoxOffset, 5, 5))
    End Sub

    Friend Sub DrawACircleInABox(ByVal g As Graphics, ByVal Box As Rectangle)

        Dim BigBoxOffset As Integer = IIf(gBoxScale = gBigBox, 5, 0)
        g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), Box)
        g.DrawEllipse(New Pen(gCurrentSkin.LineBoxes), New Rectangle(Box.X + 2 + BigBoxOffset, Box.Y + 2 + BigBoxOffset, 6, 6))
    End Sub

    Friend Sub DrawAMinusInABox(ByVal g As Graphics, ByVal Box As Rectangle)

        Dim BigBoxOffset As Integer = IIf(gBoxScale = gBigBox, 5, 0)
        g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), Box)
        g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 3 + BigBoxOffset, Box.Y + 5 + BigBoxOffset, Box.X + 7 + BigBoxOffset, Box.Y + 5 + BigBoxOffset)
    End Sub

    Friend Sub DrawAPlusInABox(ByVal g As Graphics, ByVal Box As Rectangle)

        Dim BigBoxOffset As Integer = IIf(gBoxScale = gBigBox, 5, 0)
        g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), Box)
        g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 3 + BigBoxOffset, Box.Y + 5 + BigBoxOffset, Box.X + 7 + BigBoxOffset, Box.Y + 5 + BigBoxOffset)
        g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 5 + BigBoxOffset, Box.Y + 2 + BigBoxOffset, Box.X + 5 + BigBoxOffset, Box.Y + 8 + BigBoxOffset)
    End Sub

    Friend Sub DrawASlashInABox(ByVal g As Graphics, ByVal Box As Rectangle)

        Dim BigBoxOffset As Integer = IIf(gBoxScale = gBigBox, 5, 0)
        g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), Box)
        g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 10 + BigBoxOffset, Box.Y + BigBoxOffset, Box.X + BigBoxOffset, Box.Y + 10 + BigBoxOffset)
    End Sub

    'Friend Sub DrawAnArrowABox(ByVal g As Graphics,ByVal Box As Rectangle)
    '    g.DrawRectangle(CurrentSkin.LineBoxes, Box)

    '    If RulerIsHorizontal Then
    '        If RulerHasScaleOnTopOrLeft Then
    '            'arrow down
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 5, Box.Y + 2, Box.X + 5, Box.Y + 10)
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 2, Box.Y + 7, Box.X + 5, Box.Y + 10)
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 8, Box.Y + 7, Box.X + 5, Box.Y + 10)
    '        Else
    '            'arrow up
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 5, Box.Y + 8, Box.X + 5, Box.Y)
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 2, Box.Y + 3, Box.X + 5, Box.Y)
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 8, Box.Y + 3, Box.X + 5, Box.Y)
    '        End If
    '    Else
    '        If RulerHasScaleOnTopOrLeft Then
    '            'arrow Right
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 2, Box.Y + 5, Box.X + 10, Box.Y + 5)
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 6, Box.Y + 2, Box.X + 10, Box.Y + 5)
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 6, Box.Y + 8, Box.X + 10, Box.Y + 5)
    '        Else
    '            'arrow Left
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 8, Box.Y + 5, Box.X, Box.Y + 5)
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 5, Box.Y + 2, Box.X, Box.Y + 5)
    '            g.DrawLine(CurrentSkin.LineBoxes, Box.X + 5, Box.Y + 8, Box.X, Box.Y + 5)
    '        End If
    '    End If

    'End Sub

    Friend Sub DrawANumberSignBox(ByVal g As Graphics, ByVal Box As Rectangle)
        Dim BigBoxOffset As Integer = IIf(gBoxScale = gBigBox, 5, 0)
        g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), Box)
        g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 4 + BigBoxOffset, Box.Y + 2 + BigBoxOffset, Box.X + 4 + BigBoxOffset, Box.Y + 8 + BigBoxOffset)
        g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 6 + BigBoxOffset, Box.Y + 2 + BigBoxOffset, Box.X + 6 + BigBoxOffset, Box.Y + 8 + BigBoxOffset)
        g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 2 + BigBoxOffset, Box.Y + 4 + BigBoxOffset, Box.X + 8 + BigBoxOffset, Box.Y + 4 + BigBoxOffset)
        g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 2 + BigBoxOffset, Box.Y + 6 + BigBoxOffset, Box.X + 8 + BigBoxOffset, Box.Y + 6 + BigBoxOffset)
    End Sub


    Friend Sub DrawTicksInABox(ByVal g As Graphics, ByVal Box As Rectangle, ByVal RulerIsHorizontal As Boolean, ByVal RulerHasScaleOnTopOrLeft As Boolean)

        Dim BigBoxOffset As Integer = IIf(gBoxScale = gBigBox, 5, 0)

        g.DrawRectangle(New Pen(gCurrentSkin.LineBoxes), Box)

        If RulerIsHorizontal Then
            If RulerHasScaleOnTopOrLeft Then
                'ticks down
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 2 + BigBoxOffset, Box.Y + 6 + BigBoxOffset, Box.X + 2 + BigBoxOffset, Box.Y + 10 + BigBoxOffset)
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 5 + BigBoxOffset, Box.Y + 3 + BigBoxOffset, Box.X + 5 + BigBoxOffset, Box.Y + 10 + BigBoxOffset)
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 8 + BigBoxOffset, Box.Y + 6 + BigBoxOffset, Box.X + 8 + BigBoxOffset, Box.Y + 10 + BigBoxOffset)
            Else
                'ticks up
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 2 + BigBoxOffset, Box.Y + BigBoxOffset, Box.X + 2 + BigBoxOffset, Box.Y + 4 + BigBoxOffset)
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 5 + BigBoxOffset, Box.Y + BigBoxOffset, Box.X + 5 + BigBoxOffset, Box.Y + 7 + BigBoxOffset)
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 8 + BigBoxOffset, Box.Y + BigBoxOffset, Box.X + 8 + BigBoxOffset, Box.Y + 4 + BigBoxOffset)
            End If
        Else
            If RulerHasScaleOnTopOrLeft Then
                'ticks left
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 6 + BigBoxOffset, Box.Y + 2 + BigBoxOffset, Box.X + 10 + BigBoxOffset, Box.Y + 2 + BigBoxOffset)
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 3 + BigBoxOffset, Box.Y + 5 + BigBoxOffset, Box.X + 10 + BigBoxOffset, Box.Y + 5 + BigBoxOffset)
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + 6 + BigBoxOffset, Box.Y + 8 + BigBoxOffset, Box.X + 10 + BigBoxOffset, Box.Y + 8 + BigBoxOffset)
            Else
                'ticks right
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + BigBoxOffset, Box.Y + 2 + BigBoxOffset, Box.X + 4 + BigBoxOffset, Box.Y + 2 + BigBoxOffset)
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + BigBoxOffset, Box.Y + 5 + BigBoxOffset, Box.X + 7 + BigBoxOffset, Box.Y + 5 + BigBoxOffset)
                g.DrawLine(New Pen(gCurrentSkin.LineBoxes), Box.X + BigBoxOffset, Box.Y + 8 + BigBoxOffset, Box.X + 4 + BigBoxOffset, Box.Y + 8 + BigBoxOffset)

            End If
        End If

    End Sub

    Function IsFileAvailable(ByVal filename As String, Optional ByVal WaitTimeInMilliSeconds As Integer = 0, Optional ByVal RestTimeTimeInMilliSeconds As Integer = 0) As Boolean

        Dim FileNum As Integer
        Dim FileExists As Boolean = False
        Dim ReturnCode As Boolean = False
        Dim Attribute As Microsoft.VisualBasic.FileAttribute
        Dim MaxWaitTime As DateTime = Now.AddMilliseconds(WaitTimeInMilliSeconds)
        Dim OnePass As Boolean = (WaitTimeInMilliSeconds = 0)

        Try
            While (Now < MaxWaitTime) Or OnePass
                OnePass = False

                If FileExists Then
                Else
                    Try
                        Attribute = GetAttr(filename)
                        FileExists = True
                    Catch ex As Exception
                    End Try
                End If

                If FileExists Then
                    Try
                        FileNum = FreeFile()
                        FileOpen(FileNum, filename, OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite)
                        FileClose(FileNum)
                        RestTimeTimeInMilliSeconds = 0
                        ReturnCode = True
                        Exit While
                    Catch ex As Exception
                        ReturnCode = False
                    End Try
                End If

                Application.DoEvents()
                If RestTimeTimeInMilliSeconds > 0 Then System.Threading.Thread.Sleep(RestTimeTimeInMilliSeconds)
            End While

        Catch ex As Exception
        End Try

        Return ReturnCode

    End Function


    'Do not virtualize this routine, it will fail if virtualized
    Friend Sub ShortCut(ByVal Command As String)

        ' Commands are:  
        ' Add
        ' Delete
        ' Ensure ; ensure will verify the startup link is present and correctly defined 

        Dim gThisProgramName As String = "ARuler"
        Try

            Dim MyTargetPathAndFile As String = System.Environment.CurrentDirectory & "\" & gThisProgramName
            If MyTargetPathAndFile.Contains("Rob Latour") AndAlso MyTargetPathAndFile.Contains("Debug") Then
                If Command = "Ensure" Then
                Else
                    'MsgBox("ignor request as this is a debug session")
                End If
                Exit Sub
            End If

            Dim StartupPathName As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Startup)
            Dim ShortCutFileName As String = gThisProgramName & ".lnk"
            Dim StartupPathAndShortCutFileName As String = StartupPathName & "\" & ShortCutFileName

            If Command = "Ensure" Then

                If System.IO.File.Exists(StartupPathAndShortCutFileName) Then

                    Dim obj As New Object
                    obj = CreateObject("Shell.Application")
                    Dim ShortCutsProgramPathandFileName As String = obj.NameSpace(System.Environment.SpecialFolder.Startup).ParseName(ShortCutFileName).GetLink.Path.ToString

                    If ShortCutsProgramPathandFileName.Contains(MyTargetPathAndFile) Then
                        Exit Try 'its all good
                    Else

                        Command = "Recreate"

                    End If

                Else

                    Command = "Add"

                End If

            End If


            If (Command = "Delete") OrElse (Command = "Recreate") Then

                Try

                    If System.IO.File.Exists(StartupPathName & "\" & gThisProgramName & ".lnk") Then
                        Dim fso As Object
                        fso = CreateObject("Scripting.FileSystemObject")
                        fso.DeleteFile(StartupPathName & "\" & gThisProgramName & ".lnk")
                    End If

                Catch ex As Exception
                End Try

            End If

            If Command = "Add" OrElse (Command = "Recreate") Then

                If System.IO.File.Exists(StartupPathAndShortCutFileName) Then
                    'already there, no add required
                Else
                    Try
                        Dim Shell As Object
                        Shell = CreateObject("WScript.Shell")
                        Dim MyShortcut As Object
                        Dim ProgramFilePath As String = System.Environment.CurrentDirectory
                        MyShortcut = Shell.CreateShortcut(StartupPathName & "\" & "A Ruler for Windows Quick Start.lnk")
                        MyShortcut.targetpath = Shell.ExpandEnvironmentStrings(ProgramFilePath & "\" & gThisProgramName)
                        MyShortcut.WorkingDirectory = Shell.ExpandEnvironmentStrings(ProgramFilePath)
                        MyShortcut.WindowStyle = 4
                        MyShortcut.IconLocation = Shell.ExpandEnvironmentStrings((ProgramFilePath & "\" & gThisProgramName & ".exe") & ", 0")
                        MyShortcut.Arguments = " /quickstart"
                        MyShortcut.Save()
                    Catch ex As Exception
                        '
                    End Try
                End If

            End If

        Catch ex As Exception
            '
        End Try

    End Sub

    Friend Sub SetLanguage()

        ' if the langage has not previousily been set, set it to the same language as used in the installer

        If My.Settings.Language = "not yet set" Then

            Dim Language As String = String.Empty

            Dim appPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)

            If File.Exists(appPath & "\InstallerLanguage.txt") Then

                Dim sr As StreamReader = New StreamReader(appPath & "\InstallerLanguage.txt")
                Dim line As String = sr.ReadLine()
                sr.Close()
                sr.Dispose()

                Language = line.Trim()

            End If

            If Language.Length = 2 Then
            Else
                Language = "en"
            End If

            Dim validLanguages = "ar de en es fr it nl pl pt sv"

            If validLanguages.contains(Language) Then
            Else
                Language = "en"
            End If

            Select Case Language

                Case Is = "en"

                    If System.Globalization.CultureInfo.CurrentUICulture.ToString.ToUpper.Contains("EN-GB") Then
                        Language = "en-GB"
                    Else
                        Language = "en-US"
                    End If

                Case Is = "nl"
                    Language = "nl-NL"

                Case Is = "sv",
                    Language = "sv-SE"

                Case Else
                    'Language = Language

            End Select

            My.Settings.Language = Language
            My.Settings.Save()

        End If


        ' Language resource files (*.resources) are stored in:

        ' C:\Users\Rob Latour\Documents\VBNet\eruler\eruler\eruler\Resources

        ' Set the Language Code
        Try

            If (My.Settings.Language.ToUpper = "EN-US") OrElse (My.Settings.Language.ToUpper = "EN-GB") Then
                GoTo Done
            End If

            If My.Settings.Language.ToUpper.StartsWith("AR") Then
                If My.Settings.Language.Length > 2 Then
                    My.Settings.Language = "ar"
                    My.Settings.Save()
                End If
                GoTo Done
            End If

            If My.Settings.Language.ToUpper.StartsWith("DE") Then
                If My.Settings.Language.Length > 2 Then
                    My.Settings.Language = "de"
                    My.Settings.Save()
                End If
                GoTo Done
            End If

            If My.Settings.Language.ToUpper.StartsWith("FR") Then
                If My.Settings.Language.Length > 2 Then
                    My.Settings.Language = "fr"
                    My.Settings.Save()
                End If
                GoTo Done
            End If

            If My.Settings.Language.ToUpper.StartsWith("NL") Then
                My.Settings.Language = "nl-NL"
                My.Settings.Save()
                GoTo Done
            End If

            If My.Settings.Language.ToUpper.StartsWith("PT") Then
                If My.Settings.Language.Length > 2 Then
                    My.Settings.Language = "pt"
                    My.Settings.Save()
                End If
                GoTo Done
            End If

            If My.Settings.Language.ToUpper.StartsWith("PL") Then
                If My.Settings.Language.Length > 2 Then
                    My.Settings.Language = "pl"
                    My.Settings.Save()
                End If
                GoTo Done
            End If

            If My.Settings.Language.ToUpper.StartsWith("IT") Then
                If My.Settings.Language.Length > 2 Then
                    My.Settings.Language = "it"
                    My.Settings.Save()
                End If
                GoTo Done
            End If

            If My.Settings.Language.ToUpper.StartsWith("ES") Then
                If My.Settings.Language.Length > 2 Then
                    My.Settings.Language = "es"
                    My.Settings.Save()
                End If
                GoTo Done
            End If

            If My.Settings.Language.ToUpper.StartsWith("SV") Then
                My.Settings.Language = "sv-SE"
                My.Settings.Save()
                GoTo Done
            End If

            My.Settings.Language = "en-US"
            My.Settings.Save()

Done:

            ' Set the language that is used for forms
            If My.Settings.Language = System.Globalization.CultureInfo.CurrentUICulture.ToString Then
            Else
                System.Threading.Thread.CurrentThread.CurrentUICulture = New System.Globalization.CultureInfo(My.Settings.Language)
            End If

            Dim workingCulture As String = System.Globalization.CultureInfo.CurrentUICulture.ToString

            'Set the language for use with in code literals
            gResourceManager = New ResourceManager("aruler.LanguageFile", Assembly.GetExecutingAssembly)

            '#If DEBUG Then
            '            Console.WriteLine("Current language:")
            '            Console.WriteLine(gResourceManager.GetString("0001"))
            '            Console.WriteLine(gResourceManager.GetString("0002"))
            '#End If

        Catch ex As Exception

        End Try

        SetCommonKeys()
        SetLicense()

    End Sub

    Friend Sub SetCommonKeys()

        gKey_Clear = gResourceManager.GetString("0032")
        gKey_ShowLength = gResourceManager.GetString("0033")
        gKey_ShowMidPoint = gResourceManager.GetString("0034")
        gKey_ShowThirds = gResourceManager.GetString("0035")
        gKey_ShowGoldenRatioLine = gResourceManager.GetString("0036")
        gKey_ReverseNumbers = gResourceManager.GetString("0037")
        gKey_SwitchBetweenRulerAndGuide = gResourceManager.GetString("0038")
        gKey_Home = gResourceManager.GetString("0039")
        gKey_Escape = gResourceManager.GetString("0040")
        gKey_PageUpPageDown = gResourceManager.GetString("0077")

    End Sub
    Friend Sub SetLicense()

        gLicenseWebSite = gResourceManager.GetString("0041")

    End Sub

    Friend Function FindPath(ByVal FullFileName As String) As String

        Try
            If FullFileName.Length > 0 Then
                Dim x As Integer
                For x = FullFileName.Length To 1 Step -1
                    If (Mid(FullFileName, x, 1) = "\") Or (Mid(FullFileName, x, 1) = "/") Then
                        Return Microsoft.VisualBasic.Left(FullFileName, x)
                        Exit Function
                    End If
                Next x
            End If
            Return ""
        Catch ex As Exception
            Return ""
        End Try

    End Function

    <System.Diagnostics.DebuggerStepThrough()> Friend Function QuickFilter(ByVal InputString As String, ByVal cAllowableCharacters As String) As String

        Dim OutputValue As String = String.Empty
        Try
            Dim x As Int32
            For x = 1 To Len(InputString)
                If cAllowableCharacters.IndexOf(Mid(InputString, x, 1)) = -1 Then
                Else
                    OutputValue = OutputValue & Mid(InputString, x, 1)
                End If
            Next
        Catch ex As Exception
            OutputValue = String.Empty
        End Try

        Return OutputValue

    End Function

    Public Sub RefreshAllSkinNames()

        Dim WorkingTable(5000) As String

        Dim strFileSize As String = ""
        Dim di As New IO.DirectoryInfo(XML_Path_Name)
        Dim aryFi As IO.FileInfo() = di.GetFiles("RulerDefinition_*" & MyExtention)
        Dim fi As IO.FileInfo

        Dim x As Integer = 0
        For Each fi In aryFi
            WorkingTable(x) = fi.Name.Replace("RulerDefinition_", "").Replace(MyExtention, "")
            ' WorkingTable(x) = TranslateASkinNameFromEnglishtToAnotherLanguage(WorkingTable(x))
            x += 1
        Next

        If x > 0 Then
            ReDim Preserve WorkingTable(x - 1)
            Array.Sort(WorkingTable)
        Else
            WorkingTable = Nothing
        End If

        gAllSkinNames = WorkingTable
        WorkingTable = Nothing

    End Sub

    Friend Function TranslateASkinNameFromEnglishtToAnotherLanguage(ByVal SkinName As String) As String

        'Translate a skin name to another language if it is one of the predefined skin names

        Dim ReturnCode As String = SkinName

        If My.Settings.Language.StartsWith("en") Then
        Else
            Select Case SkinName

                Case Is = "Wood"
                    ReturnCode = gResourceManager.GetString("0058")

                Case Is = "Plastic see thru"
                    ReturnCode = gResourceManager.GetString("0059")

                Case Is = "Yellow"
                    ReturnCode = gResourceManager.GetString("0060")

                Case Is = "Stainless steel"
                    ReturnCode = gResourceManager.GetString("0061")

            End Select
        End If

        Return ReturnCode

    End Function

    Friend Function TranslateASkinNameFromAnotherLanguageToEnglish(ByVal SkinName As String) As String

        'Translate a skin name to another language if it is one of the predefined skin names

        Dim ReturnCode As String = SkinName

        If My.Settings.Language.StartsWith("en") Then
        Else

            If SkinName = gResourceManager.GetString("0058") Then
                ReturnCode = "Wood"

            ElseIf SkinName = gResourceManager.GetString("0059") Then
                ReturnCode = "Plastic see thru"

            ElseIf SkinName = gResourceManager.GetString("0060") Then
                ReturnCode = "Yellow"

            ElseIf SkinName = gResourceManager.GetString("0061") Then
                ReturnCode = "Stainless steel"

            End If

        End If

        Return ReturnCode

    End Function

    Friend Sub CreateFullDirectory(ByVal DirectoryName As String)

        ' min length = 3 ; example "c:\"
        If DirectoryName.Length < 4 Then Exit Sub

        Dim x As Integer
        Dim ws As String = Microsoft.VisualBasic.Left(DirectoryName, 3)
        For x = 4 To Len(DirectoryName)
            If Mid(DirectoryName, x, 1) = "\" Then
                Try
                    If Not System.IO.Directory.Exists(ws) Then System.IO.Directory.CreateDirectory(ws)
                Catch ex As Exception
                End Try
            End If
            ws &= Mid(DirectoryName, x, 1)
        Next

        If DirectoryName.EndsWith("\") Then
        Else
            Try
                System.IO.Directory.CreateDirectory(ws)
            Catch ex As Exception
            End Try
        End If

    End Sub

    Friend Sub AdvanceToNextSkin()

        RefreshAllSkinNames()

        If gAllSkinNames.Length = 1 Then
            gNextEnabledSkinName = gCurrentSkin.Name
            Exit Sub
        End If

        Dim x As Integer = 0
        For x = 0 To gAllSkinNames.Length - 1
            If gAllSkinNames(x) = gCurrentSkin.Name Then
                Exit For
            End If
        Next

        gCurrentSkin.Name = gAllSkinNames((x + 1) Mod gAllSkinNames.Length)

        Call LoadSkin(gCurrentSkin.Name)

        My.Settings.CurrentSkinName = gCurrentSkin.Name
        My.Settings.Save()

        SetNextAndPreviousEnabledSkinNames()

    End Sub

    Friend Sub SetNextAndPreviousEnabledSkinNames()

        Dim StartingPoint As Integer
        Dim x As Integer

        Dim NumberOfSkins As Integer = gAllSkinNames.Length

        'Find the Current Skin
        For StartingPoint = 0 To NumberOfSkins - 1
            If gAllSkinNames(StartingPoint) = gCurrentSkin.Name Then
                Exit For
            End If
        Next

        'Find the next enabled skin
        x = StartingPoint
        Do
            x = (x + 1) Mod NumberOfSkins
        Loop Until GetEnabledSetting(gAllSkinNames(x))
        gNextEnabledSkinName = gAllSkinNames(x)

        'Find the previous enabled skin
        x = StartingPoint
        Do
            x = x - 1
            If x < 0 Then x = NumberOfSkins - 1
        Loop Until GetEnabledSetting(gAllSkinNames(x))
        gPreviousEnabledSkinName = gAllSkinNames(x)

    End Sub

    Friend Function GetEnabledSetting(ByVal SkinName As String) As Boolean

        SkinName = TranslateASkinNameFromAnotherLanguageToEnglish(SkinName)

        Dim ReturnCode As Boolean = False

        Try

            Dim XML_File_Name As String = XML_Path_Name & "\RulerDefinition_" & SkinName & MyExtention

            Dim ruler As Ruler = XML.ObjectXMLSerializer(Of Ruler).Load(XML_File_Name)
            ruler = XML.ObjectXMLSerializer(Of Ruler).Load(XML_File_Name)
            ReturnCode = ruler.SkinEnabled

        Catch ex As Exception

            Call MsgBox(ex.Message.ToString)

            '  Call MsgBox("Unable to retrieve the skin enabled setting with ruler definition for the following skin: '" & SkinName, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, "A Ruler for Windows - Warning")
            Call MsgBox(gResourceManager.GetString("0064") & " " & SkinName, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, gResourceManager.GetString("0052"))

        End Try

        Return ReturnCode

    End Function

    Friend Function LoadSkin(ByVal SkinName As String) As Boolean

        Dim ReturnCode As Boolean = False

        SkinName = TranslateASkinNameFromAnotherLanguageToEnglish(SkinName)

        Dim XML_File_Name As String = XML_Path_Name & "\RulerDefinition_" & SkinName & MyExtention

        Try

            If System.IO.File.Exists(XML_File_Name) Then

                Dim ruler As Ruler = XML.ObjectXMLSerializer(Of Ruler).Load(XML_File_Name)

                If ruler Is Nothing Then

                    ' Call MsgBox("Unable to load ruler definition from " & XML_File_Name, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, "A Ruler for Windows - Warning")
                    Call MsgBox(gResourceManager.GetString("0063") & " " & XML_File_Name, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, gResourceManager.GetString("0052"))

                Else

                    With gCurrentSkin

                        .Name = ruler.Name

                        .SkinEnabled = ruler.SkinEnabled

                        .LineBoxes = ruler.LineBoxes
                        .LineTicksAndSides = ruler.LineTicksAndSides

                        .Numbers = ruler.Numbers

                        .LineLength = ruler.LineLength
                        .LineMeasuring = ruler.LineMeasuring
                        .LineMidpoint = ruler.LineMidpoint

                        If ruler.LineThirds.ToArgb = 0 Then ruler.LineThirds = Color.SteelBlue
                        .LineThirds = ruler.LineThirds

                        If ruler.LineGoldenRatio.ToArgb = 0 Then ruler.LineGoldenRatio = Color.Gold
                        .LineGoldenRatio = ruler.LineGoldenRatio

                        .Tiled = ruler.Tiled

                        .HorizontalRotation = ruler.HorizontalRotation
                        .VerticalRotation = ruler.VerticalRotation

                        .TransparencyColour = ruler.TransparencyColour

                        .HorizontalBackground = New Bitmap(ruler.HorizontalBackground)
                        .VerticalBackground = New Bitmap(ruler.VerticalBackground)

                        .BackgroundImageFileName = ruler.BackgroundImageFileName

                        .Opacity = ruler.Opacity

                        gCurrentBaseBitMap = New Bitmap(.HorizontalBackground)

                        Select Case .HorizontalRotation

                            Case Is = 1

                            Case Is = 2
                                gCurrentBaseBitMap.RotateFlip(RotateFlipType.Rotate270FlipNone)

                            Case Is = 3
                                gCurrentBaseBitMap.RotateFlip(RotateFlipType.Rotate180FlipNone)

                            Case Is = 4
                                gCurrentBaseBitMap.RotateFlip(RotateFlipType.Rotate90FlipNone)

                        End Select


                    End With

                    ReturnCode = True

                End If

            Else

                'Call MsgBox("Unable to find the ruler definition for the '" & SkinName & "' skin.", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, "A Ruler for Windows - Warning"
                Call MsgBox(gResourceManager.GetString("0063") & " " & SkinName, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, gResourceManager.GetString("0052"))

            End If

        Catch ex As Exception

            Try
                System.IO.File.Delete(XML_File_Name)
            Catch ex1 As Exception
            End Try

        End Try

        Return ReturnCode

    End Function

    Friend Sub SaveCurrentSkin(ByVal SkinName As String)

        SkinName = TranslateASkinNameFromAnotherLanguageToEnglish(SkinName)

        Dim XML_File_Name As String = XML_Path_Name & "\RulerDefinition_" & SkinName & MyExtention

        Try

            Dim SaveSkin As Ruler = New Ruler

            With SaveSkin

                .Name = TranslateASkinNameFromAnotherLanguageToEnglish(gCurrentSkin.Name)
                .SkinEnabled = gCurrentSkin.SkinEnabled

                .LineBoxes = gCurrentSkin.LineBoxes
                .LineTicksAndSides = gCurrentSkin.LineTicksAndSides

                .Numbers = gCurrentSkin.Numbers

                .LineLength = gCurrentSkin.LineLength
                .LineMeasuring = gCurrentSkin.LineMeasuring
                .LineMidpoint = gCurrentSkin.LineMidpoint
                .LineThirds = gCurrentSkin.LineThirds
                .LineGoldenRatio = gCurrentSkin.LineGoldenRatio

                .TransparencyColour = gCurrentSkin.TransparencyColour

                .Tiled = gCurrentSkin.Tiled

                .HorizontalRotation = gCurrentSkin.HorizontalRotation
                .VerticalRotation = gCurrentSkin.VerticalRotation

                .HorizontalBackground = New Bitmap(gCurrentSkin.HorizontalBackground)
                .VerticalBackground = New Bitmap(gCurrentSkin.VerticalBackground)

                .BackgroundImageFileName = gCurrentSkin.BackgroundImageFileName

                .Opacity = gCurrentSkin.Opacity

            End With

            XML.ObjectXMLSerializer(Of Ruler).Save(SaveSkin, XML_File_Name)

            SaveSkin = Nothing

        Catch ex As Exception

            '  Call MsgBox("Unable to save the ruler definition for the following skin: '" & SkinName, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, "A Ruler for Windows - Warning")
            Call MsgBox(gResourceManager.GetString("0065") & " " & SkinName, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, gResourceManager.GetString("0052"))

        End Try

    End Sub

    Friend Sub ConfirmCurrentSkinHasReadWriteAuthority(ByVal XML_File_Name As String)

        Try

            Dim fi As New IO.FileInfo(XML_File_Name)
            If (fi.Attributes And IO.FileAttributes.ReadOnly) = IO.FileAttributes.ReadOnly Then
                fi.Attributes = fi.Attributes And (Not IO.FileAttributes.ReadOnly) 'Clear ReadOnly attribute
            End If

        Catch ex As Exception
        End Try

    End Sub

    Friend Sub DeleteSkin(ByVal SkinName As String)

        On Error Resume Next

        SkinName = TranslateASkinNameFromAnotherLanguageToEnglish(SkinName)

        Dim XML_File_Name = XML_Path_Name & "\RulerDefinition_" & SkinName & MyExtention
        If System.IO.File.Exists(XML_File_Name) Then System.IO.File.Delete(XML_File_Name)

    End Sub

    Friend Sub SetCurrentSkinToGeneralDefaults()

        With gCurrentSkin

            .SkinEnabled = True

            .LineBoxes = Color.Black
            .LineTicksAndSides = Color.Black
            .LineTicksAndSides = Color.Black

            .Numbers = Color.Black

            .LineLength = Color.Green
            .LineMeasuring = Color.Red
            .LineMidpoint = Color.Blue
            .LineThirds = Color.SteelBlue
            .LineGoldenRatio = Color.Gold

            .TransparencyColour = MyDefaultTransparentColour

            .Tiled = False
            .HorizontalRotation = 1
            .VerticalRotation = 1

            .Opacity = 1

            If (.Name = "Wood") OrElse (.Name = "Plastic see thru") OrElse (.Name = "Yellow") OrElse (.Name = "Stainless steel") Then

                SetDefaultBackgrounds(.Name)

            Else

                gCurrentBaseBitMap = New Bitmap(My.Resources.white)
                .HorizontalBackground = New Bitmap(My.Resources.white)
                .VerticalBackground = New Bitmap(My.Resources.white)
                .BackgroundImageFileName = "none"

            End If

        End With

    End Sub

    Friend Sub SetDefaultBackgrounds(ByVal SkinName As String)

        SkinName = TranslateASkinNameFromAnotherLanguageToEnglish(SkinName)

        With gCurrentSkin

            .SkinEnabled = True
            .Tiled = False
            .BackgroundImageFileName = "embedded"

            Select Case SkinName

                Case Is = "Wood"
                    .HorizontalBackground = New Bitmap(My.Resources.wood)
                    .VerticalRotation = 2
                    .VerticalBackground = New Bitmap(My.Resources.wood)
                    .VerticalBackground.RotateFlip(RotateFlipType.Rotate90FlipNone)

                Case Is = "Plastic see thru"
                    .HorizontalBackground = New Bitmap(My.Resources.steel)
                    .VerticalBackground = New Bitmap(My.Resources.steel)
                    .Opacity = 0.5

                Case Is = "Yellow"
                    .HorizontalBackground = New Bitmap(My.Resources.yellow)
                    .VerticalBackground = New Bitmap(My.Resources.yellow)

                Case Is = "Stainless steel"
                    .HorizontalBackground = New Bitmap(My.Resources.steel)
                    .VerticalBackground = New Bitmap(My.Resources.steel)

            End Select

        End With

    End Sub

    Friend Function IsThisPointInBounds(ByVal pt As Point) As Boolean

        ' Determines if a point is within the bounds of any screen

        Dim ConvertedPt As Point = New Point(gMinX + pt.X, gMinY + pt.Y)

        Dim ReturnCode As Boolean = False

        For ScreenNumber As Integer = 0 To gScreens.Length - 1
            If gScreens(ScreenNumber).Bounds.Contains(ConvertedPt) Then
                ReturnCode = True
                Exit For
            End If
        Next

        Return ReturnCode

    End Function

    Public Function SetImageOpacity(image As Image, opacity As Single) As Image
        Try
            'create a Bitmap the size of the image provided  
            Dim bmp As New Bitmap(image.Width, image.Height)

            'create a graphics object from the image  
            Using gfx As Graphics = Graphics.FromImage(bmp)

                'create a color matrix object  
                Dim matrix As New ColorMatrix()

                'set the opacity  
                matrix.Matrix33 = opacity

                'create image attributes  
                Dim attributes As New ImageAttributes()

                'set the color(opacity) of the image  
                attributes.SetColorMatrix(matrix, ColorMatrixFlag.[Default], ColorAdjustType.Bitmap)

                'now draw the image  
                gfx.DrawImage(image, New Rectangle(0, 0, bmp.Width, bmp.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, attributes)

            End Using

            Return bmp

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try

    End Function



    Friend Function FindAnUnusedColour(ByVal CurrentBaseBitMap As Bitmap) As Color

        Dim ReturnValue As Color = Color.FromArgb(255, 1, 1, 1)

        Dim ColourTable(255, 255, 255) As Boolean

        Try

            Dim mybitmap As Bitmap = CurrentBaseBitMap
            Dim PixelColor As Color

            'build a table of all used colours

            For w As Integer = 0 To mybitmap.Width - 1
                For h As Integer = 0 To mybitmap.Height - 1
                    PixelColor = mybitmap.GetPixel(w, h)
                    If PixelColor.A = 255 Then 'only look at alpha 255 pixels
                        ColourTable(PixelColor.R, PixelColor.G, PixelColor.B) = True
                    End If
                Next
            Next

            If ColourTable(250, 250, 250) Then

                ' If the default transparency is in use find another

                ' Find a color where r, g and b are all the same; 
                ' I'm not sure why but in testing if they were different then 
                ' oddly the primary screen would be locked out under the ruler also '
                ' I needed to start at 2,2,2 as 0,0,0 is black (too common) and 1,1,1 seems to be treated as black as well

                For x As Integer = 2 To 255
                    If ColourTable(x, x, x) Then
                    Else
                        ReturnValue = Color.FromArgb(255, x, x, x)
                        Exit For
                    End If
                Next

            End If

        Catch ex As Exception

        End Try

        ColourTable = Nothing

        Return ReturnValue

    End Function

    Friend Function IsThisColourBeingUsed(ByVal CurrentBaseBitMap As Bitmap, ByRef CurrentColour As Color) As Boolean

        Dim PixeColour As Color

        For w As Integer = 0 To CurrentBaseBitMap.Width - 1
            For h As Integer = 0 To CurrentBaseBitMap.Height - 1
                PixeColour = CurrentBaseBitMap.GetPixel(w, h)
                If PixeColour = CurrentColour Then
                    Return True ' winner winner chicken dinner
                End If
            Next
        Next

        Return False

    End Function

End Module



Public NotInheritable Class Environment

    Private Sub New()
    End Sub

End Class

<SuppressUnmanagedCodeSecurityAttribute()>
Friend NotInheritable Class SafeNativeMethods

    Private Sub New()
    End Sub

    Friend Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)

    Friend Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Integer, ByVal bInitialOwner As Integer, ByVal lpName As String) As Integer

    Friend Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short

    Friend Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" (
      ByVal process As IntPtr,
      ByVal minimumWorkingSetSize As IntPtr,
      ByVal maximumWorkingSetSize As IntPtr) As Integer

    Friend Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Integer,
    ByVal hWndInsertAfter As Integer,
    ByVal x As Integer,
    ByVal y As Integer,
    ByVal cx As Integer,
    ByVal cy As Integer,
    ByVal wFlags As Integer) As Integer

    Public Shared SRCCOPY As Integer = &HCC0020
    <DllImport("gdi32.dll")>
    Public Shared Function BitBlt(ByVal hObject As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hObjectSource As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As Integer) As Boolean
    End Function
    <DllImport("gdi32.dll")>
    Public Shared Function CreateCompatibleBitmap(ByVal hDC As IntPtr, ByVal nWidth As Integer, ByVal nHeight As Integer) As IntPtr
    End Function
    <DllImport("gdi32.dll")>
    Public Shared Function CreateCompatibleDC(ByVal hDC As IntPtr) As IntPtr
    End Function
    <DllImport("gdi32.dll")>
    Public Shared Function DeleteDC(ByVal hDC As IntPtr) As Boolean
    End Function
    <DllImport("gdi32.dll")>
    Public Shared Function DeleteObject(ByVal hObject As IntPtr) As Boolean
    End Function
    <DllImport("gdi32.dll")>
    Public Shared Function SelectObject(ByVal hDC As IntPtr, ByVal hObject As IntPtr) As IntPtr
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure 'RECT 
    <DllImport("user32.dll")>
    Public Shared Function GetDesktopWindow() As IntPtr
    End Function
    <DllImport("user32.dll")>
    Public Shared Function GetWindowDC(ByVal hWnd As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll")>
    Public Shared Function ReleaseDC(ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Integer
    End Function
    <DllImport("user32.dll")>
    Public Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef rect As RECT) As Integer
    End Function

    <DllImport("gdi32.dll")>
    Public Shared Function GetPixel(ByVal hWnd As IntPtr, X As Integer, iy As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Public Shared Function ToUnicode(ByVal virtualKeyCode As UInteger, ByVal scanCode As UInteger, ByVal keyboardState As Byte(),
    <Out, MarshalAs(UnmanagedType.LPWStr, SizeConst:=64)> ByVal receivingBuffer As StringBuilder, ByVal bufferSize As Integer, ByVal flags As UInteger) As Integer
    End Function

    <DllImport("user32.dll")> Friend Shared Function GetDpiForWindow(ByVal hWnd As IntPtr) As Integer
    End Function

End Class

