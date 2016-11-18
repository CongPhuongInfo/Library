Imports System
Imports System.Diagnostics
Imports System.Runtime
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading

Public Class ProcessWindows


    Dim mX, mY, lastX, lastY As Integer
    Dim hWnd As IntPtr
    Public Sub New()
        hWnd = IntPtr.Zero
        mX = mY = lastX = lastY = 0
    End Sub
#Region "AppActivates"
    ''' <summary>
    ''' Bắt đầu can thiệp vào ứng dụng đang chạy. Bằng cách sử dụng FindWindow
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
    ''' Bắt đầu can thiệp vào ứng dụng đang chạy. Bằng cách sử dụng GetProcessName
    ''' </summary>
    ''' <param name="Processname"></param>
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
    ''' Bắt đầu can thiệp vào ứng dụng đang chạy. Bằng cách sử dụng GetProcessById
    ''' </summary>
    ''' <param name="Id"></param>
    Public Sub AppActivate(ByVal Id As Integer)
        Dim p As Process = Process.GetProcessById(Id)
        If p.HasExited = False Then
            hWnd = p.MainWindowHandle
        End If
        SetForegroundWindow(hWnd)
    End Sub
#End Region
#Region "Mouse"
    ''' <summary>
    ''' Lấy thông tin về kích cỡ cửa sổ của ứng dụng đang bị can thiệp.
    ''' </summary>
    ''' <returns></returns>
    Public Function GetWinRectangle() As Rectangle
        Dim rect As New LPRECT()
        GetWindowRect(hWnd, rect)
        Dim mrect As New Rectangle()
        mrect.X = rect.Left
        mrect.Y = rect.Right
        mrect.Width = (rect.Right - rect.Left)
        mrect.Height = (rect.Bottom - rect.Top)
        Return mrect
    End Function
    Public Sub MoveMouseTo(ByVal x As Integer, ByVal y As Integer)
        If SetCursorPos(x, y) Then
            mX = x
            mY = y
        End If
    End Sub
    ''' <summary>
    ''' Lấy thông tin vị trí của con trỏ tại ứng dụng đang bị can thiệp. Khi có con trỏ hiển thị.
    ''' </summary>
    ''' <returns></returns>
    Public Function GetCursor() As CursorInfo
        Dim ci As CursorInfo = New CursorInfo()
        ci.cbSize = Marshal.SizeOf(ci)
        GetCursorInfo(ci)
        Return ci
    End Function
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
    Private Shared Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
        Return (HiWord << 16) Or (LoWord And &HFFFF)
    End Function

    Public Sub SendMouseLeftClick()
        Dim dword As Long = MakeDWord(mX - lastX, mY - lastY)
        SendNotifyMessage(hWnd, 516, 1, dword)
        Thread.Sleep(100)
        SendNotifyMessage(hWnd, 517, 1, dword)
    End Sub
#End Region
#Region "KeyBoard"
    ''' <summary>
    ''' Gửi sự kiện từ bàn phím tới ứng dụng đang bị can thiệp.
    ''' Sự kiện có thể là phím tắt hoặc là nội dung văn bản.
    ''' </summary>
    ''' <param name="key"></param>
    Public Sub Sendkey(ByVal key As String)
        If Not String.IsNullOrEmpty(key) Then
            SendKeys.SendWait(key)
        End If
    End Sub
#End Region
#Region "Windows API"
    'https://msdn.microsoft.com/en-us/library/windows/desktop/ff468919(v=vs.85).aspx
    Private Declare Auto Function FindWindow Lib "USER32.DLL" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Auto Function SetForegroundWindow Lib "USER32.DLL" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Auto Function GetWindowRect Lib "USER32.DLL" (ByVal hWnd As IntPtr, ByRef lpRect As LPRECT) As Boolean

    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms648396(v=vs.85).aspx
    Private Declare Auto Function ShowCursor Lib "USER32.DLL" (ByVal bShow As Boolean) As Integer
    Private Declare Auto Function SetCursorPos Lib "USER32.DLL" (ByVal x As Integer, ByVal y As Integer) As Boolean
    Private Declare Auto Function GetCursorInfo Lib "USER32.DLL" (ByRef pci As CursorInfo) As Boolean

    'https://msdn.microsoft.com/en-us/library/windows/desktop/ff468859(v=vs.85).aspx
    Private Declare Auto Function IsWindowEnabled Lib "USER32.DLL" (ByVal hWnd As IntPtr) As Boolean

    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms644947(v=vs.85).aspx  
    Private Declare Auto Function SendMessage Lib "USER32.DLL" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As UIntPtr, ByVal lParam As IntPtr) As Integer
    Private Declare Auto Function SendNotifyMessage Lib "USER32.DLL" (ByVal hWnd As IntPtr, Msg As Integer, ByVal wParam As UIntPtr, ByVal lParam As IntPtr) As Boolean
#End Region
#Region "Enums"
    Private Structure LPRECT
        Public Property Left() As Integer
            Get
                Return m_Left
            End Get
            Set
                m_Left = Value
            End Set
        End Property
        Private m_Left As Integer
        Public Property Top() As Integer
            Get
                Return m_Top
            End Get
            Set
                m_Top = Value
            End Set
        End Property
        Private m_Top As Integer
        Public Property Right() As Integer
            Get
                Return m_Right
            End Get
            Set
                m_Right = Value
            End Set
        End Property
        Private m_Right As Integer
        Public Property Bottom() As Integer
            Get
                Return m_Bottom
            End Get
            Set
                m_Bottom = Value
            End Set
        End Property
        Private m_Bottom As Integer

    End Structure
#End Region
#Region "Structs"
    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms648381(v=vs.85).aspx
    <StructLayout(LayoutKind.Sequential)>
    Public Structure CursorInfo
        Public cbSize As Int32
        Public flags As Int32
        Public hCursor As IntPtr
        Public ptScreenPos As Point
    End Structure
#End Region
End Class
