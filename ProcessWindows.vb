Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms

Public Class ProcessWindows

    Private mX, mY, lastX, lastY As Integer
    Private hWnd As IntPtr

    Public Sub New()
        hWnd = IntPtr.Zero
        mX = mY = lastX = lastY = 0
    End Sub

#Region "AppActivate"
    ''' <summary>
    ''' Bắt đầu can thiệp vào ứng dụng đang chạy bằng ClassName và WindowName.
    ''' </summary>
    Public Sub AppActivate(ByVal ClassName As String, ByVal WindowsName As String)
        hWnd = FindWindow(ClassName, WindowsName)
        If IsWindowEnabled(hWnd) = False Then
            MessageBox.Show("Không tìm thấy ứng dụng đang chạy.")
            Return
        End If
        SetForegroundWindow(hWnd)
    End Sub

    ''' <summary>
    ''' Bắt đầu can thiệp vào ứng dụng đang chạy bằng tên process.
    ''' </summary>
    Public Sub AppActivate(ByVal Processname As String)
        Dim p As Process() = Process.GetProcessesByName(Processname)
        If p.Count > 0 Then
            hWnd = p(0).MainWindowHandle
        Else
            MessageBox.Show("Không tìm thấy ứng dụng đang chạy.")
            Return
        End If
        SetForegroundWindow(hWnd)
    End Sub

    ''' <summary>
    ''' Bắt đầu can thiệp vào ứng dụng đang chạy bằng Process ID.
    ''' </summary>
    Public Sub AppActivate(ByVal Id As Integer)
        Dim p As Process = Process.GetProcessById(Id)
        If p.HasExited = False Then
            hWnd = p.MainWindowHandle
        Else
            MessageBox.Show("Không tìm thấy ứng dụng đang chạy.")
            Return
        End If
        SetForegroundWindow(hWnd)
    End Sub
#End Region

#Region "Window"
    ''' <summary>Lấy kích thước và vị trí cửa sổ đang bị can thiệp.</summary>
    Public Function GetWinRectangle() As Rectangle
        Dim rect As New LPRECT()
        GetWindowRect(hWnd, rect)
        Dim mrect As New Rectangle()
        mrect.X = rect.Left
        mrect.Y = rect.Top
        mrect.Width = (rect.Right - rect.Left)
        mrect.Height = (rect.Bottom - rect.Top)
        Return mrect
    End Function

    ''' <summary>Lấy tiêu đề (title) của cửa sổ đang bị can thiệp.</summary>
    Public Function GetWindowTitle() As String
        Dim sb As New StringBuilder(256)
        GetWindowText(hWnd, sb, sb.Capacity)
        Return sb.ToString()
    End Function

    ''' <summary>Kiểm tra cửa sổ có đang hiển thị không.</summary>
    Public Function IsWindowVisible() As Boolean
        Return IsWindowVisibleAPI(hWnd)
    End Function

    ''' <summary>Thu nhỏ cửa sổ xuống thanh taskbar.</summary>
    Public Sub MinimizeWindow()
        ShowWindow(hWnd, SW_MINIMIZE)
    End Sub

    ''' <summary>Phóng to cửa sổ toàn màn hình.</summary>
    Public Sub MaximizeWindow()
        ShowWindow(hWnd, SW_MAXIMIZE)
    End Sub

    ''' <summary>Khôi phục cửa sổ về kích thước bình thường.</summary>
    Public Sub RestoreWindow()
        ShowWindow(hWnd, SW_RESTORE)
    End Sub

    ''' <summary>
    ''' Di chuyển và/hoặc thay đổi kích thước cửa sổ.
    ''' </summary>
    Public Sub SetWindowBounds(x As Integer, y As Integer, width As Integer, height As Integer)
        MoveWindow(hWnd, x, y, width, height, True)
    End Sub

    ''' <summary>Di chuyển cửa sổ đến vị trí (x, y), giữ nguyên kích thước.</summary>
    Public Sub SetWindowPosition(x As Integer, y As Integer)
        Dim rect As Rectangle = GetWinRectangle()
        MoveWindow(hWnd, x, y, rect.Width, rect.Height, True)
    End Sub

    ''' <summary>Thay đổi kích thước cửa sổ, giữ nguyên vị trí.</summary>
    Public Sub SetWindowSize(width As Integer, height As Integer)
        Dim rect As Rectangle = GetWinRectangle()
        MoveWindow(hWnd, rect.X, rect.Y, width, height, True)
    End Sub
#End Region

#Region "Mouse"
    ''' <summary>Di chuyển chuột đến tọa độ màn hình (x, y).</summary>
    Public Sub MoveMouseTo(ByVal x As Integer, ByVal y As Integer)
        If SetCursorPos(x, y) Then
            mX = x
            mY = y
        End If
    End Sub

    ''' <summary>Lấy tọa độ chuột hiện tại trên màn hình.</summary>
    Public Function GetMousePosition() As Point
        Dim pt As New POINT()
        GetCursorPos(pt)
        Return New Point(pt.X, pt.Y)
    End Function

    ''' <summary>Lấy CursorInfo khi con trỏ đang hiển thị.</summary>
    Public Function GetCursor() As CursorInfo
        Dim ci As CursorInfo = New CursorInfo()
        ci.cbSize = Marshal.SizeOf(ci)
        GetCursorInfo(ci)
        Return ci
    End Function

    ''' <summary>
    ''' Di chuyển chuột vào góc trên-trái của cửa sổ (offset +10, +45) rồi đọc CursorInfo.
    ''' Dùng để lấy cursor khi game thay đổi con trỏ theo vị trí màn hình.
    ''' Lưu ý: hàm này có side effect (di chuyển chuột thật sự).
    ''' </summary>
    Public Function GetNoCursor() As CursorInfo
        Dim rec As Rectangle = GetWinRectangle()
        MoveMouseTo(rec.X + 10, rec.Y + 45)
        lastX = rec.X
        lastY = rec.Y
        Thread.Sleep(15)
        Dim ci As CursorInfo = New CursorInfo()
        ci.cbSize = Marshal.SizeOf(ci)
        GetCursorInfo(ci)
        Return ci
    End Function

    ''' <summary>Click chuột trái tại vị trí hiện tại của mX, mY (sau MoveMouseTo).</summary>
    Public Sub SendMouseLeftClick()
        Dim dword As Long = MakeDWord(mX - lastX, mY - lastY)
        SendNotifyMessage(hWnd, WM_LBUTTONDOWN, MK_LBUTTON, dword)
        Thread.Sleep(100)
        SendNotifyMessage(hWnd, WM_LBUTTONUP, 0, dword)
    End Sub

    ''' <summary>Click chuột phải tại vị trí hiện tại.</summary>
    Public Sub SendMouseRightClick()
        Dim dword As Long = MakeDWord(mX - lastX, mY - lastY)
        SendNotifyMessage(hWnd, WM_RBUTTONDOWN, MK_RBUTTON, dword)
        Thread.Sleep(100)
        SendNotifyMessage(hWnd, WM_RBUTTONUP, 0, dword)
    End Sub

    ''' <summary>Double click chuột trái tại vị trí hiện tại.</summary>
    Public Sub SendMouseDoubleClick()
        SendMouseLeftClick()
        Thread.Sleep(80)
        SendMouseLeftClick()
    End Sub

    ''' <summary>
    ''' Gửi WM_MOUSEMOVE trực tiếp vào cửa sổ mà không di chuyển chuột thật.
    ''' Hữu ích khi cần giả lập hover mà không ảnh hưởng vị trí chuột thật.
    ''' </summary>
    Public Sub SendMouseMove(x As Integer, y As Integer)
        Dim dword As Long = MakeDWord(x, y)
        SendNotifyMessage(hWnd, WM_MOUSEMOVE, 0, dword)
    End Sub

    Private Shared Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
        Return (HiWord << 16) Or (LoWord And &HFFFF)
    End Function
#End Region

#Region "Keyboard"
    ''' <summary>
    ''' Gửi chuỗi phím tắt hoặc văn bản tới cửa sổ đang được focus.
    ''' Ví dụ: Sendkey("{ENTER}"), Sendkey("Hello")
    ''' </summary>
    Public Sub Sendkey(ByVal key As String)
        If Not String.IsNullOrEmpty(key) Then
            SendKeys.SendWait(key)
        End If
    End Sub

    ''' <summary>
    ''' Gửi WM_KEYDOWN trực tiếp vào hWnd — không cần cửa sổ đang được focus.
    ''' Dùng Virtual Key Code (VK). Ví dụ: SendKeyDown(&H41) = phím 'A'
    ''' Xem: https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
    ''' </summary>
    Public Sub SendKeyDown(ByVal vkCode As Integer)
        PostMessage(hWnd, WM_KEYDOWN, CUInt(vkCode), GetKeyLParam(vkCode, True))
    End Sub

    ''' <summary>
    ''' Gửi WM_KEYUP trực tiếp vào hWnd — không cần cửa sổ đang được focus.
    ''' </summary>
    Public Sub SendKeyUp(ByVal vkCode As Integer)
        PostMessage(hWnd, WM_KEYUP, CUInt(vkCode), GetKeyLParam(vkCode, False))
    End Sub

    ''' <summary>
    ''' Gửi key nhấn và thả hoàn chỉnh (KeyDown + delay + KeyUp) trực tiếp vào hWnd.
    ''' Ví dụ: SendVirtualKey(Keys.Space)
    ''' </summary>
    Public Sub SendVirtualKey(ByVal vkCode As Integer, Optional delayMs As Integer = 50)
        SendKeyDown(vkCode)
        Thread.Sleep(delayMs)
        SendKeyUp(vkCode)
    End Sub

    ''' <summary>Overload tiện lợi dùng enum Keys của WinForms.</summary>
    Public Sub SendVirtualKey(ByVal key As Keys, Optional delayMs As Integer = 50)
        SendVirtualKey(CInt(key), delayMs)
    End Sub

    ''' <summary>
    ''' Tính lParam cho WM_KEYDOWN / WM_KEYUP theo chuẩn Windows.
    ''' </summary>
    Private Function GetKeyLParam(vkCode As Integer, isKeyDown As Boolean) As Long
        Dim scanCode As UInteger = MapVirtualKey(CUInt(vkCode), 0)
        If isKeyDown Then
            Return CLng((scanCode << 16) Or &H1)
        Else
            Return CLng((scanCode << 16) Or &HC0000001UI)
        End If
    End Function
#End Region

#Region "Windows API"
    ' Window
    Private Declare Auto Function FindWindow Lib "USER32.DLL" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Auto Function SetForegroundWindow Lib "USER32.DLL" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Auto Function GetWindowRect Lib "USER32.DLL" (ByVal hWnd As IntPtr, ByRef lpRect As LPRECT) As Boolean
    Private Declare Auto Function GetWindowText Lib "USER32.DLL" (ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    Private Declare Auto Function IsWindowEnabled Lib "USER32.DLL" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Auto Function IsWindowVisibleAPI Lib "USER32.DLL" Alias "IsWindowVisible" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Auto Function ShowWindow Lib "USER32.DLL" (ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    Private Declare Auto Function MoveWindow Lib "USER32.DLL" (ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean

    ' Cursor & Mouse
    Private Declare Auto Function ShowCursor Lib "USER32.DLL" (ByVal bShow As Boolean) As Integer
    Private Declare Auto Function SetCursorPos Lib "USER32.DLL" (ByVal x As Integer, ByVal y As Integer) As Boolean
    Private Declare Auto Function GetCursorPos Lib "USER32.DLL" (ByRef lpPoint As POINT) As Boolean
    Private Declare Auto Function GetCursorInfo Lib "USER32.DLL" (ByRef pci As CursorInfo) As Boolean

    ' Messages
    Private Declare Auto Function SendMessage Lib "USER32.DLL" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As UIntPtr, ByVal lParam As IntPtr) As Integer
    Private Declare Auto Function SendNotifyMessage Lib "USER32.DLL" (ByVal hWnd As IntPtr, Msg As Integer, ByVal wParam As UIntPtr, ByVal lParam As Long) As Boolean
    Private Declare Auto Function PostMessage Lib "USER32.DLL" (ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As UInteger, ByVal lParam As Long) As Boolean

    ' Keyboard
    Private Declare Function MapVirtualKey Lib "USER32.DLL" (ByVal uCode As UInteger, ByVal uMapType As UInteger) As UInteger
#End Region

#Region "Constants"
    ' ShowWindow commands
    Private Const SW_MINIMIZE As Integer = 6
    Private Const SW_MAXIMIZE As Integer = 3
    Private Const SW_RESTORE As Integer = 9

    ' Mouse messages
    Private Const WM_MOUSEMOVE As Integer = &H200
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_RBUTTONUP As Integer = &H205

    ' Mouse wParam flags
    Private Const MK_LBUTTON As UIntPtr = New UIntPtr(1)
    Private Const MK_RBUTTON As UIntPtr = New UIntPtr(2)

    ' Keyboard messages
    Private Const WM_KEYDOWN As UInteger = &H100
    Private Const WM_KEYUP As UInteger = &H101
#End Region

#Region "Structs"
    Private Structure LPRECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    Private Structure POINT
        Public X As Integer
        Public Y As Integer
    End Structure

    ' https://msdn.microsoft.com/en-us/library/windows/desktop/ms648381(v=vs.85).aspx
    <StructLayout(LayoutKind.Sequential)>
    Public Structure CursorInfo
        Public cbSize As Int32
        Public flags As Int32
        Public hCursor As IntPtr
        Public ptScreenPos As Point
    End Structure
#End Region

End Class
