Imports System
Imports System.ComponentModel
Imports System.Threading
Imports System.Reflection
Imports System.Runtime.InteropServices


Public Class KeyBoardHook

#Region "How to use ?"
    'Author : CongPhuongInfo
    'Address : Lao Cai - Vietnam
    'Product : Hook4VBNet - Copyright © 2008
    '-----------------------------------
    'Example : Keylog
    'Step 1 : Dim k new KeyBoardHook(true)
    'Step 2 : Creat a Sub KeyBoardPress(sender as object,e as KeyPressEventArgs)
    'Step 3 : AddHandler k.KeyPress, AddressOf KeyBoardPress
    'Step 4 : textbox1.text += CChar(e.KeyChar)
#End Region
#Region "Windows API"
    Private Delegate Function HookProc(ByVal nCode As Integer,
                                      ByVal wParam As Integer,
                                      ByVal lParam As IntPtr) As Integer

    'https://msdn.microsoft.com/en-us/library/windows/desktop/ff468842(v=vs.85).aspx
    Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal idHook As Integer) As Integer
    Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Integer,
                                             ByVal lpfn As HookProc,
                                             ByVal hMod As IntPtr,
                                             ByVal dwThreadId As Integer) As Integer
    Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal idHook As Integer,
                                           ByVal nCode As Integer,
                                           ByVal wParam As Integer,
                                           ByVal lParam As IntPtr) As Integer

    'https://msdn.microsoft.com/en-us/library/windows/desktop/ff468859(v=vs.85).aspx
    Private Declare Function GetKeyboardState Lib "user32.dll" (ByVal lpKeyState As Byte()) As Integer
    Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Integer) As Short
    Private Declare Function ToAscii Lib "user32.dll" (ByVal uVirtKey As Integer,
                                    ByVal uScanCode As Integer,
                                    ByVal lpKeyState As Byte(),
                                    ByVal lpChar As Byte(),
                                    ByVal uFlags As Integer) As Integer
    Private Declare Function ToUnicode Lib "user32.dll" (ByVal wVirtKey As Integer,
                                                         ByVal wScanCode As Integer,
                                                         ByVal lpKeyState As Byte(),
                                                         ByVal pwszBuff As String,
                                                         ByVal cchBuff As Integer,
                                                         ByVal wFlags As Integer) As Integer
#End Region
#Region "Enums"
    Private Enum VK As Byte

        SPACE = 2
        SHIFT = &H10
        CONTROL = &H11
        MENU = &H12
        PAUSE = &H13
        CAPITAL = &H14
        PRIOR = &H21 'Page UP
        [Next] = &H22 'Page DOWN
        [END] = &H23
        HOME = &H24

        LEFT = &H25 'LEFT ARROW key
        UP = &H26 'UP ARROW key
        RIGHT = &H27 'RIGHT ARROW key
        DOWN = &H28 'DOWN ARROW key

        NUMLOCK = &H90
        SCROLL = &H91
        ESCAPE = &H1B 'ESC Key
        SNAPSHOT = &H2C 'PRINT SCREEN key
        INSERT = &H2D
        DELETE = &H2E

        LSHIFT = &HA0 'Left SHIFT key
        RSHIFT = &HA1 'Right SHIFT key
        LCONTROL = &HA2
        RCONTROL = &H3
        LALT = &HA4
        RALT = &HA5
    End Enum
    Private Enum WH As Integer

        MSGFILTER = -1
        JOURNALRECORD = 0
        JOURNALPLAYBACK = 1
        KEYBOARD = 2
        GETMESSAGE = 3
        CALLWNDPROC = 4
        CBT = 5
        SYSMSGFILTER = 6
        MOUSE = 7
        DEBUG = 9
        SHELL = 10
        FOREGROUNDIDLE = 11
        CALLWNDPROCRET = 12
        KEYBOARD_LL = 13
        MOUSE_LL = 14

    End Enum
    Private Enum WM As Integer
        ACTIVATE = &H6
        APPCOMMAND = &H319
        KEYDOWN = &H100
        KEYUP = &H101
        [Char] = &H102
        DEADCHAR = &H103
        HOTKEY = &H312
        KILLFOCUS = &H8
        LBUTTONDBLCLK = &H203
        LBUTTONDOWN = &H201
        LBUTTONUP = &H202
        MBUTTONDBLCLK = &H209
        MBUTTONDOWN = &H207
        MBUTTONUP = &H520
        MOUSEMOVE = &H200
        MOUSEWHEEL = &H20A
        RBUTTONDBLCLK = &H206
        RBUTTONDOWN = &H204
        RBUTTONUP = &H205
        SETFOCUS = &H7
        SYSKEYDOWN = &H104
        SYSKEYUP = &H105
        SYSDEADCHAR = &H107
        UNICHAR = &H109
    End Enum
#End Region
#Region "Events"
    Private events As New EventHandlerList
    Public Custom Event KeyDown As KeyEventHandler
        AddHandler(ByVal value As KeyEventHandler)
            events.AddHandler("KeyDown", value)
        End AddHandler
        RemoveHandler(ByVal value As KeyEventHandler)
            events.RemoveHandler("KeyDown", value)
        End RemoveHandler
        RaiseEvent(ByVal sender As Object, ByVal e As KeyEventArgs)
            Dim eh As KeyEventHandler = TryCast(events("KeyDown"), KeyEventHandler)
            If eh IsNot Nothing Then eh.Invoke(sender, e)
        End RaiseEvent
    End Event
    Public Custom Event KeyPress As KeyPressEventHandler
        AddHandler(ByVal value As KeyPressEventHandler)
            events.AddHandler("KeyPress", value)
        End AddHandler
        RemoveHandler(ByVal value As KeyPressEventHandler)
            events.RemoveHandler("KeyPress", value)
        End RemoveHandler
        RaiseEvent(ByVal sender As Object, ByVal e As KeyPressEventArgs)
            Dim eh As KeyPressEventHandler = TryCast(events("KeyPress"), KeyPressEventHandler)
            If eh IsNot Nothing Then eh.Invoke(sender, e)
        End RaiseEvent
    End Event
    Public Custom Event KeyUp As KeyEventHandler
        AddHandler(ByVal value As KeyEventHandler)
            events.AddHandler("KeyUp", value)
        End AddHandler
        RemoveHandler(ByVal value As KeyEventHandler)
            events.RemoveHandler("KeyUp", value)
        End RemoveHandler
        RaiseEvent(ByVal sender As Object, ByVal e As KeyEventArgs)
            Dim eh As KeyEventHandler = TryCast(events("KeyUp"), KeyEventHandler)
            If eh IsNot Nothing Then eh.Invoke(sender, e)
        End RaiseEvent
    End Event
#End Region
#Region "Struct"
    <StructLayout(LayoutKind.Sequential)>
    Private Class KeyboardHookStruct
        Public vkCode As Integer
        Public scanCode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Class
#End Region
#Region "Hook"
    Dim HP As HookProc
    Dim ID As Integer
    Public Sub New(install As Boolean)
        ID = 0
        If install Then Start(True)
    End Sub
    Private Function HookCallBack(ncode As Integer, wparam As Integer, lparam As IntPtr) As Integer
        Dim handled As Boolean = False

        If ncode >= 0 AndAlso (events("KeyDown") IsNot Nothing OrElse events("KeyUp") IsNot Nothing OrElse events("KeyPress") IsNot Nothing) Then ' If (nCode >= 0) Then
            Dim KHS As KeyboardHookStruct = DirectCast(Marshal.PtrToStructure(lparam, GetType(KeyboardHookStruct)), KeyboardHookStruct)
            Dim control As Boolean = ((GetKeyState(VK.LCONTROL) And &H80) <> 0) OrElse ((GetKeyState(VK.RCONTROL) And &H80) <> 0)
            Dim shift As Boolean = ((GetKeyState(VK.LSHIFT) And &H80) <> 0) OrElse ((GetKeyState(VK.RSHIFT) And &H80) <> 0)
            Dim alt As Boolean = ((GetKeyState(VK.LALT) And &H80) <> 0) OrElse ((GetKeyState(VK.RALT) And &H80) <> 0)
            Dim capslock As Boolean = (GetKeyState(VK.CAPITAL) <> 0)
            Dim e1 As New KeyEventArgs(DirectCast(KHS.vkCode Or (If(control, CInt(Keys.Control), 0)) Or (If(shift, CInt(Keys.Shift), 0)) Or (If(alt, CInt(Keys.Alt), 0)), Keys))
            Select Case wparam
                Case WM.KEYDOWN, WM.SYSKEYDOWN
                    If events("KeyDown") IsNot Nothing Then
                        RaiseEvent KeyDown(Me, e1)
                        handled = handled OrElse e1.Handled
                    End If
                    Exit Select
                Case WM.KEYUP, WM.SYSKEYUP
                    If events("KeyUp") IsNot Nothing Then
                        RaiseEvent KeyUp(Me, e1)
                        handled = handled OrElse e1.Handled
                    End If
                    Exit Select
            End Select
            If wparam = WM.KEYDOWN AndAlso Not handled AndAlso Not e1.SuppressKeyPress AndAlso events("KeyPress") IsNot Nothing Then

                Dim keyState As Byte() = New Byte(255) {}
                Dim inBuffer As Byte() = New Byte(1) {}
                GetKeyboardState(keyState)

                If ToAscii(KHS.vkCode, KHS.scanCode, keyState, inBuffer, KHS.flags) = 1 Then

                    Dim key As Char = CChar(ChrW(inBuffer(0)))
                    If (capslock Xor shift) AndAlso [Char].IsLetter(key) Then
                        key = [Char].ToUpper(key)
                    End If
                    Dim e2 As New KeyPressEventArgs(key)
                    RaiseEvent KeyPress(Me, e2)

                    handled = handled OrElse e1.Handled
                End If
            End If
        End If
        If handled Then
            Return 1
        End If
        Return CallNextHookEx(ID, ncode, wparam, lparam)
    End Function
#End Region
#Region "Start/Stop"
    Public Sub Start(Optional ex As Boolean = False)
        If ID = 0 Then
            HP = New HookProc(AddressOf HookCallBack)
            ID = SetWindowsHookEx(WH.KEYBOARD_LL,
                                                HP,
                                                Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly.GetModules()(0)),
                                                0)
            If ID = 0 AndAlso ex Then
                Dim err As Integer = Marshal.GetLastWin32Error
                Throw New Exception(err)
            End If
        End If
    End Sub
    Public Sub [Stop](Optional ex As Boolean = False)
        If ID <> 0 Then
            Dim UH As Integer = UnhookWindowsHookEx(ID)
            ID = 0
            If ID = 0 AndAlso ex Then
                Dim err As Integer = Marshal.GetLastWin32Error
                Throw New Exception(err)
            End If
        End If
    End Sub
    Protected Overrides Sub Finalize()
        Try
            Me.Stop(False)
        Finally
            MyBase.Finalize()
        End Try
    End Sub


#End Region
End Class
