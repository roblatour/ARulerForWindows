Public Class frmPickAColour
    Private Sub frmPickAColour_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim Title As String = gResourceManager.GetString("0067")
        Title = Title.Remove(Title.IndexOf("-") + 2)
        Title &= gResourceManager.GetString("0080")

        Me.Text = Title

        Me.Location = gSameLocationAsSkinWindow

        Me.PictureBox1.Image = gCurrentSkin.HorizontalBackground
        Me.Size = New Size(gCurrentSkin.HorizontalBackground.Width + 10, gCurrentSkin.HorizontalBackground.Height + 35)
        Me.MaximumSize = Me.Size
        Me.MinimumSize = Me.Size

        MakeTopMost(Me.Handle.ToInt64)

        gCurrentSkin.TransparencyColour = FindAnUnusedColour(gCurrentSkin.HorizontalBackground)

        Application.DoEvents()

    End Sub


    Private Sub HorizontalOrVerticalRuler_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown

        'ref https://stackoverflow.com/questions/753132/how-to-get-the-colour-of-a-pixel-at-x-y-using-c

        Dim Handle As IntPtr = SafeNativeMethods.GetWindowDC(IntPtr.Zero)
        Dim ColourRGBInt As Integer = SafeNativeMethods.GetPixel(Handle, Cursor.Position.X, Cursor.Position.Y)
        Dim Result As Integer = SafeNativeMethods.ReleaseDC(Handle, IntPtr.Zero)

        Dim r As Integer = ColourRGBInt And &HFF
        Dim g As Integer = (ColourRGBInt And &HFF00) >> 8
        Dim b As Integer = (ColourRGBInt And &HFF0000) >> 16

        gCurrentSkin.TransparencyColour = System.Drawing.Color.FromArgb(255, r, g, b)

        Me.Close()

    End Sub


End Class