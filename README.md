# MIDI connector written in VB.NET

Presented here is a simple MIDI-in to MIDI-out connector. 
Also known as MIDI-router, MIDI-patchbay, MIDI-through, MIDI-thru, etc.

![MIDI connector](https://github.com/psitech/MIDI-connector-written-in-VB/assets/27091013/722e3bd0-1d4e-4d40-b062-ccfe6f1c60a0)

Just select an input device & an output device, and hit connect. MIDI data will then be forwarded transparently.

> [!NOTE]
> It supports standard 3-byte MIDI messages, not SYSEX messages.


In the ZIP file, you can find the complete project (.sln) and a Windows10 64-bit executable.

Below is the code. Used Visual Studio Community 2022.
```Visual Basic .NET
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar

Public Class Form1
    Private Declare Function midiInGetNumDevs Lib "winmm.dll" () As Integer
    Private Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (uDeviceID As Integer, ByRef lpCaps As MIDIINCAPS, uSize As Integer) As Integer
    Private Declare Function midiInOpen Lib "winmm.dll" (ByRef hMidiIn As IntPtr, uDeviceID As Integer, dwCallback As MidiInCallback, dwInstance As IntPtr, dwFlags As Integer) As Integer
    Private Declare Function midiInStart Lib "winmm.dll" (hMidiIn As IntPtr) As Integer
    Private Declare Function midiInStop Lib "winmm.dll" (hMidiIn As IntPtr) As Integer
    Private Declare Function midiInReset Lib "winmm.dll" (hMidiIn As IntPtr) As Integer
    Private Declare Function midiInClose Lib "winmm.dll" (hMidiIn As IntPtr) As Integer

    Private Declare Function midiOutGetNumDevs Lib "winmm.dll" () As Integer
    Private Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (uDeviceID As Integer, ByRef lpCaps As MIDIOUTCAPS, uSize As Integer) As Integer
    Private Declare Function midiOutOpen Lib "winmm.dll" (ByRef hMidiOut As IntPtr, uDeviceID As Integer, dwCallback As MidiInCallback, dwInstance As IntPtr, dwFlags As Integer) As Integer
    Private Declare Function midiOutStop Lib "winmm.dll" (hMidiOut As IntPtr) As Integer
    Private Declare Function midiOutReset Lib "winmm.dll" (hMidiOut As IntPtr) As Integer
    Private Declare Function midiOutClose Lib "winmm.dll" (hMidiOut As IntPtr) As Integer
    Private Declare Function midiOutShortMsg Lib "winmm.dll" (hMidiOut As IntPtr, dwMsg As IntPtr) As Integer

    Public Delegate Function MidiInCallback(hMidiIn As IntPtr, wMsg As UInteger, dwInstance As Integer, dwParam1 As Integer, dwParam2 As Integer) As Integer
    Public ptrCallback As New MidiInCallback(AddressOf MidiInProc)
    Public Const CALLBACK_FUNCTION As Integer = &H30000
    Public Const MIDI_IO_STATUS = &H20

    Public Structure MIDIINCAPS
        Dim wMid As Short 
        Dim wPid As Short 
        Dim vDriverVersion As Integer 
        <VBFixedString(32), MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public szPname As String 
        Dim dwSupport As Integer 
    End Structure

    Public Structure MIDIOUTCAPS
        Dim wMid As Short 
        Dim wPid As Short 
        Dim vDriverVersion As Integer 
        <VBFixedString(32), MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public szPname As String 
        Dim dwSupport As Integer 
    End Structure

    Dim hMidiIn As IntPtr
    Dim hMidiOut As IntPtr
    Dim DeviceInID As Integer
    Dim DeviceOutID As Integer
    Dim isConnected As Boolean = False
    Dim DevCnt As Integer

    Public Function MidiInProc(hMidiIn As IntPtr, wMsg As UInteger, dwInstance As Integer, dwParam1 As Integer, dwParam2 As Integer) As Integer
        midiOutShortMsg(hMidiOut, wMsg)
        Return Nothing
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Show()
        ButtonConnect.Enabled = False

        If midiInGetNumDevs() = 0 Then
            MsgBox("No MIDI devices connected")
            Application.Exit()
        End If

        Dim InCaps As New MIDIINCAPS

        For DevCnt = 0 To (midiInGetNumDevs - 1)
            midiInGetDevCaps(DevCnt, InCaps, Len(InCaps))
            ComboBoxInputDevice.Items.Add(InCaps.szPname)
        Next DevCnt

        Dim OutCaps As New MIDIOUTCAPS
        For DevCnt = 0 To (midiOutGetNumDevs - 1)
            midiOutGetDevCaps(DevCnt, OutCaps, Len(OutCaps))
            ComboBoxOutputDevice.Items.Add(OutCaps.szPname)
        Next DevCnt
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxInputDevice.SelectedIndexChanged
        DeviceInID = ComboBoxInputDevice.SelectedIndex
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxOutputDevice.SelectedIndexChanged
        DeviceOutID = ComboBoxOutputDevice.SelectedIndex
        ButtonConnect.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonConnect.Click
        If (isConnected = False) Then
            midiInOpen(hMidiIn, DeviceInID, ptrCallback, 0, CALLBACK_FUNCTION Or MIDI_IO_STATUS)
            midiOutOpen(hMidiOut, DeviceOutID, ptrCallback, 0, 0)
            midiInStart(hMidiIn)
            isConnected = True
            Label3.Visible = True
            ComboBoxInputDevice.Enabled = False
            ComboBoxOutputDevice.Enabled = False
            ButtonConnect.Text = "Disconnect"
        Else
            midiInStop(hMidiIn)
            midiInReset(hMidiIn)
            midiInClose(hMidiIn)
            midiOutReset(hMidiOut)
            midiOutClose(hMidiOut)
            isConnected = False
            Label3.Visible = False
            ComboBoxInputDevice.Enabled = True
            ComboBoxOutputDevice.Enabled = True
            ButtonConnect.Text = "Connect"
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        midiInStop(hMidiIn)
        midiInReset(hMidiIn)
        midiInClose(hMidiIn)
        midiOutReset(hMidiOut)
        midiOutClose(hMidiOut)
        Application.Exit()
    End Sub
End Class
```
