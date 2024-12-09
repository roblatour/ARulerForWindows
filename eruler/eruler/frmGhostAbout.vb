Public Class frmGhostAbout

    Private Sub frmGhost_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.BringToFront()

        Me.Opacity = gCurrentSkin.Opacity

        If gCurrentSkin.Tiled OrElse (gCurrentSkin.Name = "Wood") Then
            Me.BackgroundImageLayout = ImageLayout.Tile
            Application.DoEvents()
        Else
            Me.BackgroundImageLayout = ImageLayout.Stretch
            Application.DoEvents()
        End If

        Dim bm As Bitmap = New Bitmap(gCurrentSkin.HorizontalBackground)
        Me.BackgroundImage = bm

    End Sub

End Class