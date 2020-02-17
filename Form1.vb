Imports System.Drawing.Imaging

Public Class Form1
    Private Sub btnTakeScreenshot_Click(sender As Object, e As EventArgs)
        Application.DoEvents()
        Dim bmp As Bitmap = ScreenCap()
        bmp.Save(Application.StartupPath & "\Image\image.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
        bmp.Dispose()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Dim bmp As Bitmap = ScreenCap()

            bmp.Save(Application.StartupPath & "\screenshot.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
            bmp.Dispose()
        Catch ex As Exception
            Dim bmp As Bitmap = ScreenCap()

            bmp.Save(Application.StartupPath & "\screenshot.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
            bmp.Dispose()
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Visible = False
    End Sub
End Class
