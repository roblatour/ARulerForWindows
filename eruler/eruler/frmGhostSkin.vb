Public Class frmGhostSkin

    Private g As Graphics
    Private Loading As Boolean = True

    Private Sub frmGhost_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        MakeTopMost(Me.Handle.ToInt64)
        Loading = False

    End Sub


    Private Sub frmGhostSkin_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        Try

            If Loading Then Exit Sub

            Me.TransparencyKey = MyTransparentColour

            ' When frmSkins needs the ghost window to be redrawn
            ' it sets gRefreshGhostSkinWindow to true and changes the size of the ghost window 
            ' the code below causes the ghost window to be redrawn

            ' the various settings below are those that are changed by frmSkins

            Me.Opacity = gCurrentSkin.Opacity

            If gCurrentSkin.Tiled Then
                HorizontalRuler.BackgroundImageLayout = ImageLayout.Tile
                VerticalRuler.BackgroundImageLayout = ImageLayout.Tile
            Else
                HorizontalRuler.BackgroundImageLayout = ImageLayout.Stretch
                VerticalRuler.BackgroundImageLayout = ImageLayout.Stretch
            End If

            HorizontalRuler.BackgroundImage = New Bitmap(gCurrentSkin.HorizontalBackground)
            VerticalRuler.BackgroundImage = New Bitmap(gCurrentSkin.VerticalBackground)

            Application.DoEvents()

            frmGhostSkin_RedrawScreen()

        Catch ex As Exception

            MsgBox("Problem with frmGhostSkin_Resize " & vbCrLf & ex.Message.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "A Form Filler Error")

        End Try

    End Sub

    Private Sub frmGhostSkin_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged

        'I don't know why, but in testing if I moved the frmSkins window quickly the frmGhost window
        'minimizes, this code is a work around to restore it
        If WindowState = FormWindowState.Minimized Then
            WindowState = FormWindowState.Normal
        End If

    End Sub

    Private Sub frmGhostSkin_RedrawScreen()

        Try

            With gCurrentSkin

                Dim h, w As Integer

                'Horizontal Ruler
                g = HorizontalRuler.CreateGraphics

                h = Me.HorizontalRuler.Height - 1  '' testing here
                w = Me.HorizontalRuler.Width - 1

                frmGhostSkin_DrawBoxesOnSkinWindow(g, w, h, True)
                frmGhostSkin_ShowMeasurementsInPixels(w, h, True)

                g.DrawLine(New Pen(.LineTicksAndSides), 0, 0, 0, h)
                g.DrawLine(New Pen(.LineTicksAndSides), w, 0, w, h)
                g.DrawLine(New Pen(.LineTicksAndSides), 0, h, w, h)

                g.DrawLine(New Pen(.LineThirds), 116, 0, 116, 60)
                g.DrawString(CStr(CInt((CStr("117") * gRulerScalingFactorSetByUser) * 10) / 10).ToString, NumbersFont, New SolidBrush(.LineThirds), 39, 69)

                g.DrawLine(New Pen(.LineMidpoint), 174, 0, 174, 60)
                g.DrawString(CStr(CInt((CStr("175") * gRulerScalingFactorSetByUser) * 10) / 10).ToString, NumbersFont, New SolidBrush(.LineMidpoint), 137, 69)

                g.DrawLine(New Pen(.LineGoldenRatio), 215, 0, 215, 60)
                g.DrawString(CStr(CInt((CStr("216") * gRulerScalingFactorSetByUser) * 10) / 10).ToString, NumbersFont, New SolidBrush(.LineGoldenRatio), 200, 69)

                g.DrawLine(New Pen(.LineThirds), 232, 0, 232, 60)
                g.DrawString(CStr(CInt((CStr("233") * gRulerScalingFactorSetByUser) * 10) / 10).ToString, NumbersFont, New SolidBrush(.LineThirds), 270, 69)

                Dim wLenght As Integer = g.MeasureString((w).ToString, SmallNumbersFont).Width
                g.DrawString(CStr(CInt((w * gRulerScalingFactorSetByUser) * 10) / 10).ToString, SmallNumbersFont, New SolidBrush(.LineLength), w - wLenght, 67)



                'Vertical Ruler
                g = VerticalRuler.CreateGraphics
                h = Me.VerticalRuler.Height - 1
                w = Me.VerticalRuler.Width - 1 ' testing here

                frmGhostSkin_DrawBoxesOnSkinWindow(g, w, h, False)
                frmGhostSkin_ShowMeasurementsInPixels(w, h, False)

                g.DrawLine(New Pen(.LineTicksAndSides), 0, 0, w, 0)
                g.DrawLine(New Pen(.LineTicksAndSides), 0, h, w, h)
                g.DrawLine(New Pen(.LineTicksAndSides), w, 0, w, h)

                g.DrawLine(New Pen(.LineMeasuring), 0, 179, 60, 179)

                Dim w180 As Integer = g.MeasureString("180", NumbersFont).Width

                g.DrawString(CStr(CInt((180 * gRulerScalingFactorSetByUser) * 10) / 10).ToString, NumbersFont, New SolidBrush(.LineMeasuring), w / gUniversaleScale - w180, 152)

            End With

        Catch ex As Exception

            MsgBox("Problem with frmGhostSkin_RedrawScreen " & vbCrLf & ex.Message.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "A Form Filler Error")

        End Try

    End Sub

    Private NewRuler, AboutBox, ResizeLeftBox, ResizeRightBox, NumberBox, MinimizeNowBox, MagnifyBox, TickPositionBox, HorizontalVeriticalBox, CloseBox, BackgroundBox, MeasuringBox, MasterBoxA, MasterBoxB As Rectangle

    Private Sub frmGhostSkin_DrawBoxesOnSkinWindow(ByVal g As Drawing.Graphics, ByVal Width As Integer, ByVal Height As Integer, ByVal RulerIsHorizontal As Boolean)

        Try

            If RulerIsHorizontal Then
                Height = Height '* gUniversaleScale
            Else
                Width = Width '* gUniversaleScale
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

            A1 = New Rectangle(0, 0, 10, 10)
            B1 = New Rectangle(A1.X + 10, A1.Y, 10, 10)
            C1 = New Rectangle(A1.X, A1.Y + 10, 10, 10)
            D1 = New Rectangle(A1.X + 10, A1.Y + 10, 10, 10)
            E1 = New Rectangle(A1.X, A1.Y + 20, 10, 10)
            F1 = New Rectangle(A1.X + 10, A1.Y + 20, 10, 10)

            A2 = New Rectangle(Width - 20, 0, 10, 10)
            B2 = New Rectangle(A2.X + 10, A2.Y, 10, 10)
            C2 = New Rectangle(A2.X, A2.Y + 10, 10, 10)
            D2 = New Rectangle(A2.X + 10, A2.Y + 10, 10, 10)
            E2 = New Rectangle(A2.X, A2.Y + 20, 10, 10)
            F2 = New Rectangle(A2.X + 10, A2.Y + 20, 10, 10)

            A3 = New Rectangle(0, Height - 30, 10, 10)
            B3 = New Rectangle(A3.X + 10, A3.Y, 10, 10)
            C3 = New Rectangle(A3.X, A3.Y + 10, 10, 10)
            D3 = New Rectangle(A3.X + 10, A3.Y + 10, 10, 10)
            E3 = New Rectangle(A3.X, A3.Y + 20, 10, 10)
            F3 = New Rectangle(A3.X + 10, A3.Y + 20, 10, 10)

            A4 = New Rectangle(Width - 20, Height - 30, 10, 10)
            B4 = New Rectangle(A4.X + 10, A4.Y, 10, 10)
            C4 = New Rectangle(A4.X, A4.Y + 10, 10, 10)
            D4 = New Rectangle(A4.X + 10, A4.Y + 10, 10, 10)
            E4 = New Rectangle(A4.X, A4.Y + 20, 10, 10)
            F4 = New Rectangle(A4.X + 10, A4.Y + 20, 10, 10)

            If RulerIsHorizontal Then

                MagnifyBox = A3
                BackgroundBox = B3
                ResizeLeftBox = C3
                NumberBox = D3
                HorizontalVeriticalBox = E3
                TickPositionBox = F3

                AboutBox = C4
                ResizeRightBox = D4
                MinimizeNowBox = E4
                CloseBox = F4

                MasterBoxA = New Rectangle(A3.X, A3.Y, 20, 30)
                MasterBoxB = New Rectangle(C4.X, C4.Y, 20, 20)

            Else

                ResizeLeftBox = A2
                HorizontalVeriticalBox = B2
                NumberBox = C2
                TickPositionBox = D2
                BackgroundBox = E2
                MagnifyBox = F2

                AboutBox = C4
                MinimizeNowBox = D4
                ResizeRightBox = E4
                CloseBox = F4

                MasterBoxA = New Rectangle(A2.X, A2.Y, 20, 30)
                MasterBoxB = New Rectangle(C4.X, C4.Y, 20, 20)

            End If

            DrawAQuestionMarkInABox(g, AboutBox)
            DrawAPlusInABox(g, MagnifyBox)
            DrawANumberSignBox(g, NumberBox)
            DrawASquareInABox(g, ResizeLeftBox)
            DrawASlashInABox(g, HorizontalVeriticalBox)
            DrawTicksInABox(g, TickPositionBox, RulerIsHorizontal, True)
            DrawACircleInABox(g, BackgroundBox)
            DrawASquareInABox(g, ResizeRightBox)
            DrawAMinusInABox(g, MinimizeNowBox)
            DrawAnXInABox(g, CloseBox)

        Catch ex As Exception
        End Try

    End Sub

    Private Sub frmGhostSkin_ShowMeasurementsInPixels(ByVal Width As Integer, ByVal Height As Integer, ByVal rulerIsHorizontal As Boolean)

        Try

            Static Dim Increment, NumbersWritenAt As Integer

            Dim z1 As Integer = 0
            Dim z2 As Integer = 0
            Dim z3 As Integer = 0
            Dim y1 As Integer

            ' Draw Ticks

            Dim TickLocations() As Integer = {5, 10, 50, 100}

            Dim BarHeight As Integer

            BarHeight = 15

            y1 = 0
            Increment = 20
            NumbersWritenAt = 40

            Dim StartAt, EndAt, Direction As Single
            Direction = 1

            For Depth As Short = 0 To 3

                Dim StepValue As Integer = TickLocations(Depth) * Direction

                If rulerIsHorizontal Then

                    StartAt = -1
                    EndAt = Width - 1

                    For x As Single = StartAt To EndAt Step StepValue

                        If Depth > 0 Then
                            If x = 0 Then z1 = 0 Else z1 = -1
                            If x = Width Then z2 = 0 Else z2 = 1
                        End If

                        For z As Integer = z1 To z2
                            For z3 = 0 To 0
                                g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), (x + z) + z3, y1, (x + z) + z3, y1 + BarHeight)
                            Next
                        Next

                    Next

                Else

                    StartAt = -1
                    EndAt = Height - 1

                    For x As Single = StartAt To EndAt Step StepValue

                        If Depth > 0 Then
                            If x = 0 Then z1 = 0 Else z1 = -1
                            If x = Height Then z2 = 0 Else z2 = 1
                        End If
                        For z As Integer = z1 To z2
                            For z3 = 0 To 0
                                g.DrawLine(New Pen(gCurrentSkin.LineTicksAndSides), y1, (x + z) + z3, y1 + BarHeight, (x + z) + z3)
                            Next
                        Next

                    Next

                End If

                y1 += BarHeight
            Next

            'Draw Numbers
            Dim s As String
            Dim w, h, offset, inc As Short
            Dim LastWholeNumber As Integer

            If rulerIsHorizontal Then

                LastWholeNumber = (Math.Truncate(Width / 100) * 100)

                StartAt = 0
                EndAt = LastWholeNumber
                inc = 0

                For x As Short = StartAt To EndAt Step 100 * Direction

                    s = CStr(CInt((CStr(inc) * gRulerScalingFactorSetByUser) * 10) / 10)
                    inc += 100

                    w = g.MeasureString(s, NumbersFont).Width
                    offset = 0

                    g.DrawString(s, NumbersFont, New SolidBrush(gCurrentSkin.Numbers), (x + offset) - w - 0.5, NumbersWritenAt)

                Next

            Else

                LastWholeNumber = (Math.Truncate(Height / 100) * 100)

                ' skip writing the bottom number when boxes are in the way
                If (Height - LastWholeNumber < 28) Then LastWholeNumber -= 100

                StartAt = 100
                EndAt = LastWholeNumber
                inc = 100

                For x As Short = StartAt To EndAt Step 100 * Direction

                    s = CStr(CInt((CStr(inc) * gRulerScalingFactorSetByUser) * 10) / 10)
                    inc += 100

                    w = g.MeasureString(s, NumbersFont).Width
                    h = g.MeasureString(s, NumbersFont).Height / 2

                    offset = 0

                    g.DrawString(s, NumbersFont, New SolidBrush(gCurrentSkin.Numbers), Width / gUniversaleScale - w, x - h)

                Next

            End If

        Catch ex As Exception

            MsgBox("Problem with frmGhostSkin_ShowMeasurementsInPixels " & vbCrLf & ex.Message.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "A Form Filler Error")

        End Try

    End Sub

End Class