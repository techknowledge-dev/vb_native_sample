' Native Class Sample code for VB.NET
' -----------------------------------
' Copyright 2002-2022, TechKnowledge inc.
'
'
Imports System
Imports System.Text
Imports System.Text.Encoding
Imports System.Runtime.InteropServices
Imports BtLib

Public Class frmNativeSample
  Inherits System.Windows.Forms.Form

#Region " Windows form desinger generated code "

  Public Sub New()
    MyBase.New()
    InitializeComponent()
  End Sub

  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  Private components As System.ComponentModel.IContainer

  Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
  Friend WithEvents btnStart As System.Windows.Forms.Button
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.ListBox1 = New System.Windows.Forms.ListBox()
    Me.btnStart = New System.Windows.Forms.Button()
    Me.SuspendLayout()
    '
    'ListBox1
    '
    Me.ListBox1.HorizontalScrollbar = True
    Me.ListBox1.ItemHeight = 12
    Me.ListBox1.Location = New System.Drawing.Point(8, 8)
    Me.ListBox1.Name = "ListBox1"
    Me.ListBox1.ScrollAlwaysVisible = True
    Me.ListBox1.Size = New System.Drawing.Size(240, 304)
    Me.ListBox1.TabIndex = 0
    '
    'btnStart
    '
    Me.btnStart.Location = New System.Drawing.Point(256, 16)
    Me.btnStart.Name = "btnStart"
    Me.btnStart.Size = New System.Drawing.Size(96, 24)
    Me.btnStart.TabIndex = 1
    Me.btnStart.Text = "Start !"
    '
    'frmNativeSample
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
    Me.ClientSize = New System.Drawing.Size(360, 325)
    Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnStart, Me.ListBox1})
    Me.Name = "frmNativeSample"
    Me.Text = "Native Sample"
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

  End Sub

  Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click

    Static posblk(128) As Byte
    Static data(334) As Byte
    Static room As New Room()
    Static clientid As New ClientID()

    Dim dataLength As UInt16
    Dim keyBuf(128) As Byte
    Dim fname() As Byte
    Dim keyNum As Int16
    Dim rc As Int16
    Dim keyBufLen As Int16
    'Dim ptr As System.IntPtr
    Dim i As Int16
    Dim cidbuff As Byte() = New Byte(15) {}

    keyBufLen = 128

    fname = Encoding.ASCII.GetBytes("btrv:///demodata?table=room")
    data(0) = 0
    dataLength = 0

    ' open...
    rc = Native.BtrCallId(Operation.Open,
                        posblk,
                        data,
                        dataLength,
                        fname,
                        fname.Length,
                        keyNum,
                        cidbuff)

    dataLength = Marshal.SizeOf(room)

    keyNum = 0
    rc = Native.BtrCallId(Operation.GetFirst,
                        posblk,
                        data,
                        dataLength,
                        keyBuf,
                        keyBufLen,
                        keyNum,
                        cidbuff)

    ListBox1.Items.Clear()

    While rc = 0
      arrayToRoom(data, room)
      ListBox1.Items.Add(room.Number.ToString() + vbTab + room.Building_Name)
      '// get next.
      rc = Native.BtrCallId(Operation.GetNext,
                          posblk,
                          data,
                          dataLength,
                          keyBuf,
                          keyBufLen,
                          keyNum,
                          cidbuff)
    End While
    '	close!
    rc = Native.BtrCallId(Operation.Close, posblk, data, dataLength, keyBuf, keyBufLen, 0, cidbuff)
  End Sub

  Private Sub arrayToRoom(array As Byte(), r As Room)
    Dim pinned As GCHandle
    Dim ptr As IntPtr
    pinned = GCHandle.Alloc(array, GCHandleType.Pinned)
    ptr = pinned.AddrOfPinnedObject
    Marshal.PtrToStructure(ptr, r)
    pinned.Free()
  End Sub

  Private Sub roomToArray(array As Byte(), r As Room)
    Dim pinned As GCHandle
    Dim ptr As IntPtr
    pinned = GCHandle.Alloc(array, GCHandleType.Pinned)
    ptr = pinned.AddrOfPinnedObject
    Marshal.StructureToPtr(r, ptr, True)
    pinned.Free()
  End Sub

End Class


<StructLayout(LayoutKind.Sequential, pack:=1, CharSet:=CharSet.Ansi)>
Public Class Room
    Public null_flag1 As Byte
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=25)>
    Public Building_Name As String
    Public null_flag2 As Byte
    Public Number As Int32
    Public null_flag3 As Byte
    Public Capacity As Int16
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=20)>
    Public Type As String
End Class

<StructLayout(LayoutKind.Sequential, Pack:=1, CharSet:=CharSet.Ansi)>
Public Class ClientID
    Public resv(12) As Byte
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=2)>
    Public sa As String
    Public id As Int16
End Class


