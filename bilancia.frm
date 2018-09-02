VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form mf 
   Caption         =   "Bilancia20"
   ClientHeight    =   8235
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11085
   Icon            =   "bilancia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fr_settings 
      Caption         =   "Settings"
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   11055
      Begin VB.TextBox tb_setpoint 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   8
         ToolTipText     =   "enter the target value"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox tb_lastrec 
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   " - "
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cb_com 
         Height          =   315
         ItemData        =   "bilancia.frx":0442
         Left            =   600
         List            =   "bilancia.frx":0455
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "select com port"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl_setpoint 
         Caption         =   "Set Point"
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbl_lastoutput 
         Caption         =   "Last Output"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl_port 
         Caption         =   "Port"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer readTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10320
      Top             =   120
   End
   Begin MSCommLib.MSComm Comm 
      Left            =   9480
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.TextBox tb_log 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Height          =   2000
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5700
      Width           =   10575
   End
   Begin VB.Frame fr_log 
      Caption         =   "log"
      Height          =   2400
      Left            =   0
      TabIndex        =   1
      Top             =   5400
      Width           =   11055
   End
   Begin MSChart20Lib.MSChart MSC 
      Height          =   4935
      Left            =   0
      OleObjectBlob   =   "bilancia.frx":0477
      TabIndex        =   9
      Top             =   960
      Width           =   11055
   End
   Begin VB.Menu mf_menu 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "mf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public logfile As Integer
Public measfile As Integer
Dim measflushcnt As Integer
Public inBuf As String
Dim SetPoint As Single
Dim first_time As Date

Option Explicit
'
Const LOG_ERROR = 1
Const LOG_INFO = 2
Const LOG_DEBUG = 3
Const LOG_LEVEL = LOG_INFO

Private Sub cb_com_Click()
If Comm.PortOpen Then Comm.PortOpen = False
port_init
End Sub

Private Sub Comm_OnComm()
    Dim Msg As String, ErrMsg As String
    
    Msg = ""
    ErrMsg = ""
    
   Select Case Comm.CommEvent
   ' Errors
      Case comEventBreak   ' A Break was received.
      ErrMsg = "break"
      Case comEventFrame   ' Framing Error
      ErrMsg = "framing error"
      Case comEventOverrun   ' Data Lost.
      ErrMsg = "overrun"
      Case comEventRxOver   ' Receive buffer overflow.
      ErrMsg = "RX overflow"
      Case comEventRxParity   ' Parity Error.
      ErrMsg = "parity error"
      Case comEventTxFull   ' Transmit buffer full.
      ErrMsg = "TX overflow"
      Case comEventDCB   ' Unexpected error retrieving DCB]
      ErrMsg = "DBC error"

   ' Events
      Case comEvCD   ' Change in the CD line.
      Msg = "CD event"
      Case comEvCTS   ' Change in the CTS line.
      Msg = "CTS event"
      Case comEvDSR   ' Change in the DSR line.
      Msg = "DSR event"
      Case comEvRing   ' Change in the Ring Indicator.
      Msg = "Ring event"
'      Case comEvReceive   ' Received RThreshold # of chars.
'      Msg = "RX threshold event"
      Case comEvSend   ' There are SThreshold number of characters in the transmit buffer.
      Msg = "TX threshold event"
      Case comEvEOF   ' An EOF charater was found in the input stream
      Msg = "EOF event"
   End Select
   If Len(ErrMsg) > 0 Then
        Call logmsg2(LOG_ERROR, ErrMsg)
   End If
   If Len(Msg) > 0 Then
        Call logmsg2(LOG_DEBUG, Msg)
   End If
End Sub
      
Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call logmsg2(LOG_INFO, "starting version " & App.Major & "." & App.Minor & "." & App.Revision)
Call init_chart
cb_com.ListIndex = 0 'this will cause port_init to be called
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim Msg   ' Declare variable.
      Msg = "Do you really want to exit?"
   ' If user clicks the No button, stop QueryUnload.
   If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Cancel = True
End Sub

Private Sub Form_Resize()
fr_log.Left = 0
fr_log.Width = mf.Width - 100
tb_log.Width = fr_log.Width - 200
fr_log.Top = 5400
tb_log.Top = 5700

fr_settings.Left = 0
fr_settings.Width = mf.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
readTimer.Enabled = False
Call logmsg2(LOG_INFO, "exiting...")
End Sub

Sub port_init()
Dim port As Integer
tb_lastrec.Text = "- - -"
If cb_com.ListIndex < 0 Then cb_com.ListIndex = 0
port = cb_com.ItemData(cb_com.ListIndex)
readTimer.Enabled = False

On Error GoTo errhandler
Comm.CommPort = port
Comm.Settings = "9600,n,8,1"
Comm.InputLen = 0
Call logmsg2(LOG_INFO, "opening port " + CStr(Comm.CommPort) + ", Handshaking= " + CStr(Comm.Handshaking) + ", settings " + Comm.Settings)
Comm.PortOpen = True
readTimer.Enabled = True
Exit Sub
errhandler:
Call logmsg2(LOG_ERROR, "Error: " & Err.Source & " : " & Err.Description)
Exit Sub
End Sub

Sub logmsg2(l As Integer, m As String)
    If l <= LOG_LEVEL Then
        logmsg (m)
    End If
End Sub

Sub logmsg(m As String)
Dim logpath$, pre$
If logfile = 0 Then
    logpath$ = "bil20.log"
    logfile = 1
    Open logpath$ For Append As logfile
End If

   pre$ = Time()
   Print #logfile, pre$ & " " & m
   Debug.Print m
   mf.tb_log.Text = Left$(pre$ & " " & m & vbNewLine & mf.tb_log.Text, 1000)

End Sub

Private Sub readTimer_Timer()
Static inbuffer As String
Dim inString, c, char_sep, sep_str As String
Dim l, i, sep, idx As Integer

sep_str = vbCr & vbLf

' Call logmsg2(LOG_DEBUG, "timer")
readTimer.Enabled = False
If Not Comm.PortOpen Then Exit Sub
'inString = Comm.Input
inString = "123gr." & vbCr & vbLf

l = Len(inString)
If l > 0 Then
    Call logmsg2(LOG_DEBUG, "read " & CStr(l) & " characters: " & inString)
    For i = 1 To l
        c = Mid(inString, i, 1)
        Call logmsg2(LOG_DEBUG, "char " & i & "= " & c & ", ascii " & Asc(c))
    Next
End If

inbuffer = inbuffer & inString
Do
sep = 0
For idx = 1 To Len(inbuffer)
    char_sep = Mid(inbuffer, idx, 1)
    If InStr(sep_str, char_sep) > 0 Then
        sep = idx
        Exit For
    End If
Next idx

If sep > 0 Then
    process_rec_str (Left$(inbuffer, sep - 1))
    inbuffer = Mid$(inbuffer, sep + 1, Len(inbuffer) - sep)
End If

For idx = 1 To Len(inbuffer)
    char_sep = Mid$(inbuffer, idx, 1)
    If InStr(sep_str, char_sep) = 0 Then
        Exit For
    End If
Next idx
inbuffer = Mid$(inbuffer, idx, Len(inbuffer))

Loop Until sep = 0

readTimer.Enabled = True
End Sub

Sub process_rec_str(rec As String)
Dim timestamp As String
Dim meas_time As Date
Dim chdot As String
Dim reading As Single
Dim pos As Integer
Dim diff_time As Integer
Dim row As Integer
Dim meas_tmp As Single, setpoint_tmp As Single, rowlabel_tmp As String

meas_time = Now
If first_time = 0 Then
first_time = meas_time
End If
timestamp = Format(meas_time, "Long Time")
diff_time = Round((meas_time - first_time) * (24# * 60 * 60))

Call logmsg2(LOG_DEBUG, "process_record: " & timestamp & " " & rec)

chdot = Mid(1.2, 2, 1)

'strip last three characters ("gr.")
If Right$(rec, 3) = "gr." Then
    rec = Left$(rec, Len(rec) - 3)
End If

' cut ","
pos = 0
Do
    pos = InStr(rec, ",")
    If pos > 0 Then
        rec = Mid(rec, 1, pos - 1) & Mid(rec, pos + 1, Len(rec) - pos)
    Else
        Exit Do
    End If
Loop

'turn . into decimal point
Do
    pos = InStr(rec, ".")
    If pos > 0 Then
        Mid(rec, pos, 1) = chdot
    Else
        Exit Do
    End If
Loop

On Error GoTo errhandler
reading = CSng(rec)
On Error GoTo 0

tb_lastrec.Text = timestamp & " (" & diff_time & " sec)" & " - " & reading

Call write_measure(timestamp, reading)
'MSC.Repaint = False
MSC.Column = 1
MSC.RowLabel = timestamp
MSC.Data = reading

MSC.Column = 2
MSC.Data = SetPoint

If MSC.row >= MSC.RowCount Then
'shift chart left
For row = 2 To MSC.RowCount
MSC.row = row
rowlabel_tmp = MSC.RowLabel
MSC.Column = 1
meas_tmp = MSC.Data
MSC.Column = 2
setpoint_tmp = MSC.Data
MSC.row = row - 1
MSC.RowLabel = rowlabel_tmp
MSC.Column = 1
MSC.Data = meas_tmp
MSC.Column = 2
MSC.Data = setpoint_tmp
Next row
MSC.row = MSC.RowCount
Else
    MSC.row = MSC.row + 1
End If
'MSC.Repaint = True

Exit Sub

errhandler:
Call logmsg2(LOG_ERROR, "Error: " & Err.Source & " : " & Err.Description & "; rec=" & rec)
rec = ""
Exit Sub
End Sub

Sub write_measure(t As String, r As Single)
Dim measpath$
Dim line$
Dim sep$

sep$ = ";"

If measfile = 0 Then
    measpath$ = "bil20.csv"
    measfile = 2
    Open measpath$ For Append As measfile
    Print #measfile, "Timestamp" & sep$ & "reading"
    measflushcnt = 0
End If

If measflushcnt > 60 Then
' flush
    Close #measfile
    Open measpath$ For Append As measfile
    measflushcnt = 0
End If
    
line$ = t & sep$ & r
Print #measfile, line$
   
End Sub

Private Sub tb_setpoint_Change()
    Dim sp As Single
    On Error GoTo errhandler
    If tb_setpoint.Text = "" Then
    sp = 0
    Else
    sp = tb_setpoint.Text
    End If
    tb_setpoint.BackColor = vbWindowBackground
    SetPoint = sp
    Exit Sub
errhandler:
    sp = 0
    tb_setpoint.BackColor = vbRed
End Sub

Sub init_chart()
Dim x As Integer, z As Integer
For x = 1 To MSC.RowCount
MSC.row = x
MSC.RowLabel = ""
For z = 1 To 2
MSC.Column = z
MSC.Data = 0
Next z
Next x
MSC.row = 1

End Sub
