VERSION 5.00
Begin VB.Form frmPicker 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Grap screen color"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Picker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   141
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      Caption         =   "Accept"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Close"
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   8
      Top             =   1740
      Width           =   795
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Index           =   2
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Index           =   0
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4320
      Picture         =   "Picker.frx":08CA
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      ToolTipText     =   " Use the pipette to select screen color "
      Top             =   1440
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   1425
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3600
      Top             =   1560
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Output format of color code:"
      Height          =   195
      Left            =   1800
      TabIndex        =   9
      ToolTipText     =   " Select output format "
      Top             =   60
      Width           =   2655
   End
End
Attribute VB_Name = "frmPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
        
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Sub color_Implement()

On Error GoTo errhandler
    
    ' reset interval and adjustment settings
    frmMain.cmdReset_Click 0
    frmMain.cmdReset_Click 1
    
    ' disable synchronising function
    frmMain.chkLock(0).Value = 0
    frmMain.chkLock(1).Value = 0
    frmMain.chkLock(2).Value = 0
    
    frmMain.txtNewHexValue.Text = Text(1).Text
    frmMain.txtNewHexValue_KeyUp 0, False
    
errhandler:
    Exit Sub
End Sub
Private Sub cmdAction_Click(Index As Integer)

On Error GoTo errhandler

    Select Case Index
        Case 0
            color_Implement
            
        Case 1
            Unload Me
            
    End Select
    

errhandler:
    Exit Sub
End Sub

Private Sub Form_Load()

On Error GoTo errhandler
    
    
    form_StayOnTop frmPicker, "ABSOLUTE", "C"
    
    Timer1.Interval = 50
    Timer1.Enabled = False
    Check(2).Value = vbChecked

errhandler:
    Exit Sub
End Sub

Private Sub Check_Click(Index As Integer)
 
On Error GoTo errhandler

Dim i As Integer
  
    Select Case Index
        Case 0
            If Check(Index).Value = vbChecked Then
                Check(1).Value = vbUnchecked
                Check(2).Value = vbUnchecked
                Check(Index).Caption = "Copy"
            Else
                Check(Index).Caption = ""
            End If
    
        Case 1
            If Check(Index).Value = vbChecked Then
                Check(0).Value = vbUnchecked
                Check(2).Value = vbUnchecked
                Check(Index).Caption = "Copy"
            Else
                Check(Index).Caption = ""
            End If
    
        Case 2
            If Check(Index).Value = vbChecked Then
                Check(0).Value = vbUnchecked
                Check(1).Value = vbUnchecked
                Check(Index).Caption = "Copy"
            Else
                Check(Index).Caption = ""
            End If
  End Select

  For i = 0 To 2
        If Check(i).Value = vbChecked Then
            Clipboard.Clear
            Clipboard.SetText Text(i).Text
        End If
  Next i

errhandler:
    Exit Sub
End Sub

Private Sub Timer1_Timer()

On Error GoTo errhandler

Dim hWndp               As Long
Dim hDCp                As Long
Dim Result              As Long
Dim Pt                  As POINTAPI

Dim r                   As Byte
Dim g                   As Byte
Dim b                   As Byte
Dim i                   As Integer
  
Static LastX As Long
Static LastY As Long

    Call GetCursorPos(Pt)
    If Pt.X = LastX And Pt.Y = LastY Then Exit Sub
      
    ' correct for icon size (field size = 32px, icon size = 17px)
    Pt.X = Pt.X - 17
    Pt.Y = Pt.Y + 15
          
    LastX = Pt.X
    LastY = Pt.Y
      
    hWndp = WindowFromPoint(Pt.X, Pt.Y)
  
    hDCp = GetDC(hWndp)
  
    Call ScreenToClient(hWndp, Pt)
  
    Result = GetPixel(hDCp, Pt.X, Pt.Y)
    If Result = -1 Then
        Call BitBlt(Picture1.hDC, 0, 0, 1, 1, hDCp, Pt.X, Pt.Y, vbSrcCopy)
        Result = Picture1.Point(0, 0)
    End If
  
    Call ReleaseDC(hWndp, hDCp)
    If Result = -1 Then Exit Sub
    Me.Picture1.BackColor = Result
  
    Call RGBsplit(Result, r, g, b)

    Text(0).Text = Result
    Text(1).Text = CheckNull(Hex(r)) & CheckNull(Hex(g)) & CheckNull(Hex(b))
    Text(2).Text = "RGB(" & Format$(r, "000") & ", " & Format$(g, "000") & ", " & Format$(b, "000") & ")"
  
    For i = 0 To 2
        If Check(i).Value = vbChecked Then
            Clipboard.Clear
            Clipboard.SetText Text(i).Text
        End If
    Next i

errhandler:
    Exit Sub
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo errhandler

    PipetteOn

errhandler:
    Exit Sub
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
On Error GoTo errhandler

  PipetteOff
  
errhandler:
    Exit Sub
End Sub

Private Sub PipetteOn()

On Error GoTo errhandler

  Picture2.Visible = False
  frmPicker.MouseIcon = LoadPicture(App.Path & "\ColorPicker.ico") 'Picture2.Picture
  frmPicker.MousePointer = 99
  Screen.MousePointer = 99
  Timer1.Enabled = True
  
errhandler:
    Exit Sub
End Sub

Private Sub PipetteOff()

On Error GoTo errhandler
    
  Me.Timer1.Enabled = False
  Me.MousePointer = vbNormal
  Screen.MousePointer = vbNormal
  Me.Picture2.Visible = True

errhandler:
    Exit Sub
End Sub

Private Sub RGBsplit(ByVal Col, r As Byte, g As Byte, b As Byte)

On Error GoTo errhandler

  b = (Col And 16711680) / 65536
  g = (Col And 65280) / 256
  r = Col And 255

errhandler:
    Exit Sub
End Sub

Private Function CheckNull(f As String)

On Error GoTo errhandler

  If Len(f) < 2 Then f = "0" & f
  CheckNull = f

errhandler:
    Exit Function
End Function




