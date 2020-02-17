Imports System.Runtime.InteropServices

Module ScreenCapture
    <StructLayout(LayoutKind.Sequential)>
    Private Structure POINTAPI
        Public x As Int32
        Public y As Int32
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure CURSORINFO
        Public cbSize As Int32
        Public flags As Int32
        Public hCursor As IntPtr
        Public ptScreenPos As POINTAPI
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure ICONINFO
        Public fIcon As Boolean
        Public xHotspot As Int32
        Public yHotspot As Int32
        Public hbmMask As IntPtr
        Public hbmColor As IntPtr
    End Structure

    <DllImport("user32.dll", EntryPoint:="GetCursorInfo")>
    Private Function GetCursorInfo(ByRef pci As CURSORINFO) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function DrawIcon(ByVal hDC As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal hIcon As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", EntryPoint:="GetIconInfo")>
    Private Function GetIconInfo(ByVal hIcon As IntPtr, ByRef piconinfo As ICONINFO) As Boolean
    End Function

    Private Const vbSRCCOPY As Integer = &HCC0020
    Private Const CAPTUREBLT As Integer = &H40000000 ' capture layered windows

    <DllImport("gdi32.dll")>
    Private Function BitBlt(ByVal hdc As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As Integer) As Boolean
    End Function

    <DllImport("gdi32.dll")> _
    Private Function DeleteObject(ByVal hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Public Function ScreenCap(Optional IncludeMouse As Boolean = True) As Bitmap
        ' ***
        ' * If you want only the primary screen...
        ' ***
        ' Dim bmp As New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)


        ' ***
        ' * If you want ALL screens
        ' ***
        Dim bmp As Bitmap = New Bitmap(
            Screen.AllScreens.Sum(Function(s As Screen) s.Bounds.Width),
            Screen.AllScreens.Max(Function(s As Screen) s.Bounds.Height)
        )

        Dim gdest As Graphics = Graphics.FromImage(bmp)
        Dim hDestDC As IntPtr = gdest.GetHdc()
        Dim gsrc As Graphics = Graphics.FromHwnd(IntPtr.Zero)
        Dim hSrcDC As IntPtr = gsrc.GetHdc()

        BitBlt(hDestDC, 0, 0, bmp.Width, bmp.Height, hSrcDC, 0, 0, vbSRCCOPY Or CAPTUREBLT)

        If IncludeMouse Then
            Dim pcin As New CURSORINFO()
            pcin.cbSize = Marshal.SizeOf(pcin)

            If GetCursorInfo(pcin) Then
                Dim piinfo As ICONINFO

                If GetIconInfo(pcin.hCursor, piinfo) Then
                    DrawIcon(hDestDC, pcin.ptScreenPos.x - piinfo.xHotspot, pcin.ptScreenPos.y - piinfo.yHotspot, pcin.hCursor)
                    If CBool(piinfo.hbmMask) Then DeleteObject(piinfo.hbmMask)
                    If CBool(piinfo.hbmColor) Then DeleteObject(piinfo.hbmColor)
                End If
            End If
        End If

        gdest.ReleaseHdc()
        gdest.Dispose()
        gsrc.ReleaseHdc()
        gsrc.Dispose()

        Return bmp
    End Function

End Module