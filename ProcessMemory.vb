Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Text

Public Class ProcessMemory

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
#End Region
#Region "Process"
    'https://msdn.microsoft.com/en-us/library/windows/desktop/ms684880(v=vs.85).aspx
    Const PROCESS_WM_READ As Integer = &H10
    Const PROCESS_VM_WRITE As Integer = &H20
    Const PROCESS_ALL_ACCESS As Integer = &H1F0FFF
    Dim ProcessHandle As IntPtr

    Public Sub New(name As String)
        Try
            Dim p As Process = Process.GetProcessesByName(name)(0)
            ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, CUInt(p.Id))
        Catch ex As Exception
        End Try
    End Sub
    Public Sub New(id As Integer)
        Try
            ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, CUInt(id))
        Catch ex As Exception
        End Try
    End Sub

    Public Function ReadString(Position As Integer, Optional processSize As Integer = 24) As String
        Return Encoding.UTF8.GetString(ReadBytes(Position, processSize))
    End Function
    Public Function ReadInt(ByVal Position As UInteger, Optional processSize As Integer = 24) As Integer
        Dim bytes As Byte() = New Byte(processSize - 1) {}
        ReadProcessMemory(ProcessHandle, IntPtr.op_Explicit(Position), bytes, UIntPtr.op_Explicit(4), 0)
        Return BitConverter.ToInt32(bytes, 0)
    End Function
    Public Function ReadInt(ByVal Position As UInteger, ByVal offset As UInteger, Optional processSize As Integer = 24) As Integer
        Dim bytes As Byte() = New Byte(processSize - 1) {}
        Dim Address As UInteger = CUInt(ReadInt(Position)) + offset
        ReadProcessMemory(ProcessHandle, IntPtr.op_Explicit(Address), bytes, UIntPtr.op_Explicit(4), 0)
        Return BitConverter.ToInt32(bytes, 0)
    End Function
    Public Function ReadBytes(ByVal Position As Integer, ByVal processSize As UInteger) As Byte()
        Dim buffer As Byte() = New Byte(processSize - 1) {}
        ReadProcessMemory(ProcessHandle, IntPtr.op_Explicit(Position), buffer, UIntPtr.op_Explicit(processSize), 0)
        Return buffer
    End Function
    Public Sub WriteBytes(ByVal Position As Integer, ByVal buffer As Byte())
        WriteProcessMemory(ProcessHandle, Position, buffer, buffer.Length, 0)
    End Sub
    Public Sub WriteBytes(ByVal Position As Integer, ByVal buffer As String)
        Dim data As Byte() = Encoding.UTF8.GetBytes(buffer)
        WriteProcessMemory(ProcessHandle, Position, data, data.Length, 0)
    End Sub
    ''' <summary>
    ''' Change processSize --> GetObjectSize("Hello Worlds !")
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Public Function GetObjectSize(value As Object) As Integer
        Dim bf As New BinaryFormatter()
        Dim ms As New MemoryStream()
        Dim Array As Byte()
        bf.Serialize(ms, value)
        Array = ms.ToArray()
        Return Array.Length
    End Function
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
                                                                    ByVal nSize As UIntPtr,
                                                                    ByVal lpNumberOfBytesWritten As UInteger) As Boolean
#End Region
End Class
