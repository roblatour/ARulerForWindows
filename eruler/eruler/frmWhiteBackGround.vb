Public Class frmWhiteBackGround

    Private Const ScrollBarWidth As Integer = 17

    Private WidthLessScrollBars As Integer = Screen.PrimaryScreen.Bounds.Width - ScrollBarWidth
    Private HeightLessScrollBars As Integer = Screen.PrimaryScreen.Bounds.Height - ScrollBarWidth

    Private Sub frmWhiteBackGround_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.Size = ResizeMe(Me.Size) 'v3.1

        Me.Location = New Point(0, 0)
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height

        Me.TransparencyKey = MyTransparentColour
        Me.BackColor = MyTransparentColour

        Me.BackgroundImage = DrawFilledRectangle(WidthLessScrollBars, HeightLessScrollBars)

        Application.DoEvents()
        System.Threading.Thread.Sleep(1000)

        Me.BringToFront()

    End Sub


    Private Function DrawFilledRectangle(x As Integer, y As Integer) As Bitmap
        Dim bmp As New Bitmap(x, y)
        Using graph As Graphics = Graphics.FromImage(bmp)
            Dim ImageSize As New Rectangle(0, 0, x, y)
            graph.FillRectangle(Brushes.White, ImageSize)
        End Using
        Return bmp
    End Function

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Select Case e.KeyCode

            Case Is = Keys.Escape
                frmRuler.RestoreNormalScreen_Public()
                Application.DoEvents()

        End Select

    End Sub

    Private Sub frmWhiteBackGround_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        frmRuler.BringToFront()
    End Sub
    Private Sub frmWhiteBackGround_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        frmRuler.BringToFront()
    End Sub

    Dim obj1 As Object = New Object
    Private Sub frmWhiteBackGround_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        SyncLock obj1

            Try

                Me.TransparencyKey = MyTransparentColour
                Me.BackColor = MyTransparentColour

                Dim ScrollBarBackgroundColour As Color = Color.Blue
                If MyTransparentColour = ScrollBarBackgroundColour Then
                    ScrollBarBackgroundColour = Color.White
                End If

                e.Graphics.FillRectangle(New SolidBrush(ScrollBarBackgroundColour), New Rectangle(Me.Width - ScrollBarWidth, 0, ScrollBarWidth, Me.Height))
                e.Graphics.FillRectangle(New SolidBrush(ScrollBarBackgroundColour), New Rectangle(0, Me.Height - ScrollBarWidth, Me.Width, ScrollBarWidth))

            Catch ex As Exception
            End Try

        End SyncLock

    End Sub

End Class