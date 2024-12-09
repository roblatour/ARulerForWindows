Imports System
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Imaging

Namespace ScreenShotLogic

    Friend Class ScreenCapture

        Public Function CaptureScreen() As Image

            Dim img As Image = CaptureWindow(SafeNativeMethods.GetDesktopWindow())
            Application.DoEvents() ' seems necissary - don't remove
            Return img

        End Function 'CaptureScreen 

        Public Function CaptureWindow(ByVal handle As IntPtr) As Image

            Dim Dummy As Integer
            Dim hdcSrc As IntPtr = SafeNativeMethods.GetWindowDC(handle)
            Dim windowRect As New SafeNativeMethods.RECT
            Dummy = SafeNativeMethods.GetWindowRect(handle, windowRect)
            Dim width As Integer = windowRect.right - windowRect.left

            Dim height As Integer = windowRect.bottom - windowRect.top
            Dim hdcDest As IntPtr = SafeNativeMethods.CreateCompatibleDC(hdcSrc)

            Dim hBitmap As IntPtr = SafeNativeMethods.CreateCompatibleBitmap(hdcSrc, width, height)
            Dim hOld As IntPtr = SafeNativeMethods.SelectObject(hdcDest, hBitmap)
            SafeNativeMethods.BitBlt(hdcDest, 0, 0, width, height, hdcSrc, 0, 0, SafeNativeMethods.SRCCOPY)
            SafeNativeMethods.SelectObject(hdcDest, hOld)
            SafeNativeMethods.DeleteDC(hdcDest)
            Dummy = SafeNativeMethods.ReleaseDC(handle, hdcSrc)
            Dim img As Image = Image.FromHbitmap(hBitmap)
            SafeNativeMethods.DeleteObject(hBitmap)

            Return img

        End Function

        'Public Sub CaptureWindowToFile(ByVal handle As IntPtr, ByVal filename As String, ByVal format As ImageFormat)
        '    Dim img As Image = CaptureWindow(handle)
        '    img.Save(filename, format)
        'End Sub

        'Public Sub CaptureScreenToFile(ByVal filename As String, ByVal format As ImageFormat)
        '    Dim img As Image = CaptureScreen()
        '    img.Save(filename, format)
        'End Sub

    End Class
End Namespace

'Module modScreenCapture

'    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

'        Dim sc As New ScreenShotLogic.ScreenCapture
'        Dim img As Image = sc.CaptureScreen
'        Me.PictureBox1.Image = img
'        sc.CaptureWindowToFile(Me.PictureBox1.Handle, "C:\\temp2.gif", ImageFormat.Gif)

'    End Sub


'End Module