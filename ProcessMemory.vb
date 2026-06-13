Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text

Public Class ProcessMemory
    Implements IDisposable

#Region "How to use"
    'Read/Write Value From ProcessMemory of Games, Apps, ...
    'Author : CongPhuongInfo © 2008. Lao Cai - Vietnam
    'Step 1 : Dim p as New ProcessMemory("notepad") Or Id in Taskmgr
    'Step 2 : open notepad and write "Hello"
    'Step 3 : use tool "Hxd - Hexeditor" --> OpenRAM and choice "notepad" running.
    'Step 4 : find position by key word "Hello" with option "unicode string, case.., all, datatype = text-string"
    'Step 5 : Value of Position = Address (Column Left) + Offset (Row Header)
    'Step 6 : print value as string or integer
    ' console.writeline(p.readstring(&Hxxxxx))
    ' console.writeline(p.readint(&Hxxxxx))
    ' console.readline()
    '
    ' Multi-level pointer example:
    ' console.writeline(p.ReadPointerChain(&H400000, {&H10, &H2C, &H4}))
    '
    ' GetBaseAddress example:
    ' Dim base As UInteger = p.GetBaseAddress()
    ' console.writeline(p.ReadInt(base + &H1234))
#End Region

#Region "Constants"
    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms684880(v=vs.85).aspx
    Const PROCESS_VM_READ As Integer = &H10
    Const PROCESS_VM_WRITE As Integer = &H20
    Const PROCESS_ALL_ACCESS As Integer = &H1F0FFF
    Const STILL_ACTIVE As Integer = 259
#End Region

#Region "Fields"
    Private _processHandle As IntPtr
    Private _processId As Integer
    Private _processName As String
    Private _disposed As Boolean = False

    ''' <summary>Trả về True nếu handle hợp lệ và process vẫn còn chạy.</summary>
    Public ReadOnly Property IsOpen() As Boolean
        Get
            If _processHandle = IntPtr.Zero Then Return False
            Dim exitCode As Integer = 0
            GetExitCodeProcess(_processHandle, exitCode)
            Return exitCode = STILL_ACTIVE
        End Get
    End Property

    Public ReadOnly Property ProcessId() As Integer
        Get
            Return _processId
        End Get
    End Property

    Public ReadOnly Property ProcessName() As String
        Get
            Return _processName
        End Get
    End Property
#End Region

#Region "Constructors"
    Public Sub New(name As String)
        Try
            Dim p As Process = Process.GetProcessesByName(name)(0)
            _processName = name
            _processId = p.Id
            _processHandle = OpenProcess(PROCESS_ALL_ACCESS, False, CUInt(p.Id))
        Catch ex As Exception
        End Try
    End Sub

    Public Sub New(id As Integer)
        Try
            Dim p As Process = Process.GetProcessById(id)
            _processName = p.ProcessName
            _processId = id
            _processHandle = OpenProcess(PROCESS_ALL_ACCESS, False, CUInt(id))
        Catch ex As Exception
        End Try
    End Sub
#End Region

#Region "Process Info"
    ''' <summary>
    ''' Lấy địa chỉ base của module chính (EXE). Cần thiết khi game dùng ASLR.
    ''' Ví dụ: Dim base = p.GetBaseAddress() : p.ReadInt(base + &H1234)
    ''' </summary>
    Public Function GetBaseAddress() As UInteger
        Try
            Dim p As Process = Process.GetProcessById(_processId)
            Return CUInt(p.MainModule.BaseAddress.ToInt32())
        Catch ex As Exception
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Lấy địa chỉ base của một module DLL cụ thể trong process.
    ''' Ví dụ: p.GetModuleBaseAddress("GameCore.dll")
    ''' </summary>
    Public Function GetModuleBaseAddress(moduleName As String) As UInteger
        Try
            Dim p As Process = Process.GetProcessById(_processId)
            For Each m As ProcessModule In p.Modules
                If m.ModuleName.ToLower() = moduleName.ToLower() Then
                    Return CUInt(m.BaseAddress.ToInt32())
                End If
            Next
        Catch ex As Exception
        End Try
        Return 0
    End Function
#End Region

#Region "Read"
    ''' <summary>
    ''' Đọc chuỗi từ bộ nhớ. Mặc định dùng Unicode (UTF-16 LE) phù hợp game Hàn/Trung.
    ''' Truyền enc:=Encoding.UTF8 nếu game dùng UTF-8.
    ''' </summary>
    Public Function ReadString(Position As UInteger, Optional size As Integer = 64, Optional enc As Encoding = Nothing) As String
        If enc Is Nothing Then enc = Encoding.Unicode
        Dim raw As Byte() = ReadBytes(Position, CUInt(size))
        ' Cắt tại null terminator nếu có
        Dim nullPos As Integer = -1
        Dim step As Integer = If(enc.Equals(Encoding.Unicode), 2, 1)
        Dim i As Integer = 0
        While i < raw.Length - (step - 1)
            Dim isNull As Boolean = True
            For j As Integer = 0 To step - 1
                If raw(i + j) <> 0 Then isNull = False : Exit For
            Next
            If isNull Then nullPos = i : Exit While
            i += step
        End While
        If nullPos >= 0 Then
            Return enc.GetString(raw, 0, nullPos)
        End If
        Return enc.GetString(raw)
    End Function

    ''' <summary>Đọc 4 bytes trả về Int32.</summary>
    Public Function ReadInt(ByVal Position As UInteger) As Integer
        Dim bytes As Byte() = New Byte(3) {}
        ReadProcessMemory(_processHandle, IntPtr.op_Explicit(Position), bytes, UIntPtr.op_Explicit(4), 0)
        Return BitConverter.ToInt32(bytes, 0)
    End Function

    ''' <summary>Đọc pointer rồi cộng offset, trả về Int32 tại địa chỉ kết quả.</summary>
    Public Function ReadInt(ByVal Position As UInteger, ByVal offset As UInteger) As Integer
        Dim address As UInteger = CUInt(ReadInt(Position)) + offset
        Return ReadInt(address)
    End Function

    ''' <summary>Đọc 4 bytes trả về UInt32.</summary>
    Public Function ReadUInt(ByVal Position As UInteger) As UInteger
        Dim bytes As Byte() = New Byte(3) {}
        ReadProcessMemory(_processHandle, IntPtr.op_Explicit(Position), bytes, UIntPtr.op_Explicit(4), 0)
        Return BitConverter.ToUInt32(bytes, 0)
    End Function

    ''' <summary>Đọc 8 bytes trả về Int64.</summary>
    Public Function ReadLong(ByVal Position As UInteger) As Long
        Dim bytes As Byte() = New Byte(7) {}
        ReadProcessMemory(_processHandle, IntPtr.op_Explicit(Position), bytes, UIntPtr.op_Explicit(8), 0)
        Return BitConverter.ToInt64(bytes, 0)
    End Function

    ''' <summary>Đọc 2 bytes trả về Int16.</summary>
    Public Function ReadShort(ByVal Position As UInteger) As Short
        Dim bytes As Byte() = New Byte(1) {}
        ReadProcessMemory(_processHandle, IntPtr.op_Explicit(Position), bytes, UIntPtr.op_Explicit(2), 0)
        Return BitConverter.ToInt16(bytes, 0)
    End Function

    ''' <summary>Đọc 1 byte trả về Byte.</summary>
    Public Function ReadByte(ByVal Position As UInteger) As Byte
        Dim bytes As Byte() = New Byte(0) {}
        ReadProcessMemory(_processHandle, IntPtr.op_Explicit(Position), bytes, UIntPtr.op_Explicit(1), 0)
        Return bytes(0)
    End Function

    ''' <summary>Đọc 4 bytes trả về Single (float).</summary>
    Public Function ReadFloat(ByVal Position As UInteger) As Single
        Dim bytes As Byte() = New Byte(3) {}
        ReadProcessMemory(_processHandle, IntPtr.op_Explicit(Position), bytes, UIntPtr.op_Explicit(4), 0)
        Return BitConverter.ToSingle(bytes, 0)
    End Function

    ''' <summary>Đọc 8 bytes trả về Double.</summary>
    Public Function ReadDouble(ByVal Position As UInteger) As Double
        Dim bytes As Byte() = New Byte(7) {}
        ReadProcessMemory(_processHandle, IntPtr.op_Explicit(Position), bytes, UIntPtr.op_Explicit(8), 0)
        Return BitConverter.ToDouble(bytes, 0)
    End Function

    ''' <summary>Đọc 1 byte trả về Boolean (0 = False, khác 0 = True).</summary>
    Public Function ReadBool(ByVal Position As UInteger) As Boolean
        Return ReadByte(Position) <> 0
    End Function

    ''' <summary>Đọc n bytes thô từ bộ nhớ.</summary>
    Public Function ReadBytes(ByVal Position As UInteger, ByVal size As UInteger) As Byte()
        Dim buffer As Byte() = New Byte(size - 1) {}
        ReadProcessMemory(_processHandle, IntPtr.op_Explicit(Position), buffer, UIntPtr.op_Explicit(size), 0)
        Return buffer
    End Function

    ''' <summary>
    ''' Đọc multi-level pointer chain. Rất phổ biến trong game hack.
    ''' Ví dụ: ReadPointerChain(&H400000, {&H10, &H2C, &H4})
    ''' Tương đương: [[[base + &H10] + &H2C] + &H4]
    ''' </summary>
    Public Function ReadPointerChain(baseAddress As UInteger, offsets As UInteger()) As Integer
        Dim address As UInteger = baseAddress
        For i As Integer = 0 To offsets.Length - 2
            address = CUInt(ReadInt(address)) + offsets(i)
        Next
        Return ReadInt(address + offsets(offsets.Length - 1))
    End Function

    ''' <summary>Tương tự ReadPointerChain nhưng trả về Float.</summary>
    Public Function ReadPointerChainFloat(baseAddress As UInteger, offsets As UInteger()) As Single
        Dim address As UInteger = baseAddress
        For i As Integer = 0 To offsets.Length - 2
            address = CUInt(ReadInt(address)) + offsets(i)
        Next
        Return ReadFloat(address + offsets(offsets.Length - 1))
    End Function
#End Region

#Region "Write"
    ''' <summary>Ghi Int32 vào bộ nhớ.</summary>
    Public Sub WriteInt(ByVal Position As UInteger, ByVal value As Integer)
        WriteBytes(Position, BitConverter.GetBytes(value))
    End Sub

    ''' <summary>Ghi UInt32 vào bộ nhớ.</summary>
    Public Sub WriteUInt(ByVal Position As UInteger, ByVal value As UInteger)
        WriteBytes(Position, BitConverter.GetBytes(value))
    End Sub

    ''' <summary>Ghi Int64 vào bộ nhớ.</summary>
    Public Sub WriteLong(ByVal Position As UInteger, ByVal value As Long)
        WriteBytes(Position, BitConverter.GetBytes(value))
    End Sub

    ''' <summary>Ghi Int16 vào bộ nhớ.</summary>
    Public Sub WriteShort(ByVal Position As UInteger, ByVal value As Short)
        WriteBytes(Position, BitConverter.GetBytes(value))
    End Sub

    ''' <summary>Ghi 1 byte vào bộ nhớ.</summary>
    Public Sub WriteByte(ByVal Position As UInteger, ByVal value As Byte)
        WriteBytes(Position, New Byte() {value})
    End Sub

    ''' <summary>Ghi Single (float) vào bộ nhớ.</summary>
    Public Sub WriteFloat(ByVal Position As UInteger, ByVal value As Single)
        WriteBytes(Position, BitConverter.GetBytes(value))
    End Sub

    ''' <summary>Ghi Double vào bộ nhớ.</summary>
    Public Sub WriteDouble(ByVal Position As UInteger, ByVal value As Double)
        WriteBytes(Position, BitConverter.GetBytes(value))
    End Sub

    ''' <summary>Ghi Boolean vào bộ nhớ (True = 1, False = 0).</summary>
    Public Sub WriteBool(ByVal Position As UInteger, ByVal value As Boolean)
        WriteByte(Position, If(value, CByte(1), CByte(0)))
    End Sub

    ''' <summary>Ghi chuỗi vào bộ nhớ. Mặc định UTF-16 LE, nhất quán với ReadString.</summary>
    Public Sub WriteString(ByVal Position As UInteger, ByVal value As String, Optional enc As Encoding = Nothing)
        If enc Is Nothing Then enc = Encoding.Unicode
        WriteBytes(Position, enc.GetBytes(value))
    End Sub

    ''' <summary>Ghi mảng bytes thô vào bộ nhớ.</summary>
    Public Sub WriteBytes(ByVal Position As UInteger, ByVal buffer As Byte())
        WriteProcessMemory(_processHandle, IntPtr.op_Explicit(Position), buffer, CUInt(buffer.Length), 0)
    End Sub
#End Region

#Region "Utilities"
    ''' <summary>
    ''' Lấy kích thước unmanaged của một value type/struct.
    ''' Dùng để xác định size khi đọc/ghi.
    ''' </summary>
    Public Function GetObjectSize(Of T As Structure)(value As T) As Integer
        Return Marshal.SizeOf(value)
    End Function
#End Region

#Region "IDisposable"
    ''' <summary>
    ''' Giải phóng process handle. Gọi khi không còn dùng nữa.
    ''' Hoặc dùng Using block: Using p = New ProcessMemory("game") ... End Using
    ''' </summary>
    Public Sub Close()
        Dispose()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If Not _disposed Then
            If _processHandle <> IntPtr.Zero Then
                CloseHandle(_processHandle)
                _processHandle = IntPtr.Zero
            End If
            _disposed = True
        End If
    End Sub
#End Region

#Region "Windows API"
    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms684320(v=vs.85).aspx
    Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As UInteger,
                                                             ByVal bInheritHandle As Boolean,
                                                             ByVal dwProcessId As UInteger) As IntPtr
    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms680553(v=vs.85).aspx
    Private Declare Function ReadProcessMemory Lib "Kernel32.dll" (ByVal hProcess As IntPtr,
                                                                   ByVal lpBaseAddress As IntPtr,
                                                                   ByVal lpBuffer As Byte(),
                                                                   ByVal nSize As UIntPtr,
                                                                   ByVal lpNumberOfBytesRead As UInteger) As Boolean
    Private Declare Function WriteProcessMemory Lib "Kernel32.dll" (ByVal hProcess As IntPtr,
                                                                    ByVal lpBaseAddress As IntPtr,
                                                                    ByVal lpBuffer As Byte(),
                                                                    ByVal nSize As UInteger,
                                                                    ByVal lpNumberOfBytesWritten As UInteger) As Boolean
    Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As IntPtr) As Boolean
    Private Declare Function GetExitCodeProcess Lib "Kernel32.dll" (ByVal hProcess As IntPtr,
                                                                    ByRef lpExitCode As Integer) As Boolean
#End Region

End Class
