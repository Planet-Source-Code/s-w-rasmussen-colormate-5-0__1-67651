VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmColor 
   BackColor       =   &H00E9F3F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Color Picker"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   ControlBox      =   0   'False
   FillColor       =   &H00E9F3F4&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808080&
   Icon            =   "Color.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   309
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   210
      Left            =   120
      Max             =   64
      TabIndex        =   21
      Top             =   3765
      Width           =   2055
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   135
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   20
      ToolTipText     =   " Complementary "
      Top             =   3195
      Width           =   2025
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4080
      Picture         =   "Color.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00E9F3F4&
      Caption         =   "Exit"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4950
      Width           =   585
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00E9F3F4&
      Caption         =   "Minimize"
      Height          =   285
      Index           =   2
      Left            =   3735
      TabIndex        =   17
      Top             =   4950
      Width           =   795
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   135
      ScaleHeight     =   510
      ScaleWidth      =   2025
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   " Previous "
      Top             =   360
      Width           =   2025
   End
   Begin VB.Timer Timer1 
      Left            =   4020
      Top             =   720
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00E9F3F4&
      Caption         =   "Reset Previous"
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4950
      Width           =   1305
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00E9F3F4&
      Caption         =   "Open Editor"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4950
      Width           =   1065
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3675
      Index           =   2
      LargeChange     =   10
      Left            =   2280
      Max             =   0
      Min             =   255
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   300
      Width           =   210
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00E9F3F4&
      Caption         =   "Store"
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   " left: Selected - right: Complement - shift: Previous  "
      Top             =   4950
      Width           =   600
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3675
      Index           =   1
      LargeChange     =   10
      Left            =   2580
      Max             =   0
      Min             =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   210
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3675
      Index           =   0
      LargeChange     =   10
      Left            =   2880
      Max             =   0
      Min             =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   210
   End
   Begin RichTextLib.RichTextBox txtPos 
      Height          =   225
      Index           =   2
      Left            =   3780
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"Color.frx":1194
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4185
      Picture         =   "Color.frx":120F
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   " Use pipette to pick screen color "
      Top             =   4500
      Width           =   240
   End
   Begin RichTextLib.RichTextBox txtPos 
      Height          =   225
      Index           =   0
      Left            =   1500
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4380
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"Color.frx":1799
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtPos 
      Height          =   225
      Index           =   4
      Left            =   1500
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4620
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"Color.frx":1814
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt1DecR 
      Height          =   225
      Left            =   2295
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4380
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Color.frx":188F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt1DecG 
      Height          =   225
      Left            =   2745
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4380
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Color.frx":190A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt1DecB 
      Height          =   225
      Left            =   3195
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4380
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Color.frx":1985
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt4DecR 
      Height          =   225
      Left            =   2295
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4620
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Color.frx":1A00
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt4DecG 
      Height          =   225
      Left            =   2745
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4620
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Color.frx":1A7B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt4DecB 
      Height          =   225
      Left            =   3195
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4620
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Color.frx":1AF6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtPos 
      Height          =   225
      Index           =   3
      Left            =   1500
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4140
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"Color.frx":1B71
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt3DecR 
      Height          =   225
      Left            =   2295
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4140
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Color.frx":1BEC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt3DecG 
      Height          =   225
      Left            =   2745
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4140
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Color.frx":1C67
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt3DecB 
      Height          =   225
      Left            =   3195
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4140
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   397
      _Version        =   393217
      BackColor       =   16777215
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Color.frx":1CE2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   " Selected "
      Top             =   345
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E9F3F4&
      Caption         =   "Previous:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   180
      TabIndex        =   36
      Top             =   4155
      UseMnemonic     =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E9F3F4&
      Caption         =   "Selected:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   180
      TabIndex        =   31
      Top             =   4395
      UseMnemonic     =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E9F3F4&
      Caption         =   "Complmt:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   180
      TabIndex        =   30
      Top             =   4635
      UseMnemonic     =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E9F3F4&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1290
      TabIndex        =   16
      Top             =   90
      Width           =   45
   End
   Begin VB.Label lblY 
      BackColor       =   &H00E9F3F4&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1425
      TabIndex        =   15
      Top             =   90
      Width           =   600
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E9F3F4&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   600
      TabIndex        =   14
      Top             =   90
      Width           =   600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E9F3F4&
      Caption         =   "X-Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   13
      Top             =   90
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      Height          =   450
      Left            =   4080
      Shape           =   1  'Square
      Top             =   4395
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E9F3F4&
      Height          =   3855
      Left            =   3180
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   2940
      TabIndex        =   8
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   2340
      TabIndex        =   6
      Top             =   60
      Width           =   135
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const WU_LOGPIXELSX = 88
Const WU_LOGPIXELSY = 90

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private C()                 As Long

Private LockGradient_M      As Boolean
Private PosX_M              As Long
Private PosY_M              As Long

Function ConvertTwipsToPixels(lngTwips As Single, lngDirection As Single) As Long

On Error GoTo errhandler

Dim lngDC                   As Long
Dim lngPixelsPerInch        As Long

Const nTwipsPerInch = 1440
   
   lngDC = GetDC(0)
   
   If (lngDirection = 0) Then                                   ' Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
      
   Else                                                         ' Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
      
   End If
   
   lngDC = ReleaseDC(0, lngDC)
   
   ConvertTwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch

errhandler:
    Exit Function
End Function
Private Function CheckNull(f As String)

On Error GoTo errhandler

  If Len(f) < 2 Then f = "0" & f
  CheckNull = f

errhandler:
    Exit Function
End Function

Private Sub RGBSplit_Complement(ByVal Col, r As Integer, g As Integer, b As Integer)

On Error GoTo errhandler
        
    b = 255 - (Col And 16711680) / 65536
    g = 255 - (Col And 65280) / 256
    r = 255 - Col And 255
    
errhandler:
    Exit Sub
End Sub

Private Sub PipetteOff()

On Error GoTo errhandler
    
Dim r As Integer
Dim g As Integer
Dim b As Integer
    
    lblX.Caption = PosX_M
    lblY.Caption = PosY_M
        
    Me.Timer1.Enabled = False
    Me.MousePointer = vbNormal
    Screen.MousePointer = vbNormal
    Me.Picture2.Visible = True
    
    Call RGBsplit(Picture1.BackColor, r, g, b)
    
    Picture4.BackColor = RGB(255 - r, 255 - g, 255 - b)
    
    HScroll1.Value = 0
    HScroll1_Scroll
    
    Me.Visible = True
    
errhandler:
    Exit Sub
End Sub

Private Sub PipetteOn(Optional ByVal HideWhilePipetting As Boolean = False)

On Error GoTo errhandler
    
    lblX.Caption = 0
    lblY.Caption = 0
    
    Picture2.Visible = False
    frmColor.MouseIcon = Picture5.Picture
    frmColor.MousePointer = 99
    Screen.MousePointer = 99
    HScroll1.Value = 0
    Timer1.Enabled = True
    
    If HideWhilePipetting Then
        Me.Visible = False
    End If
    
errhandler:
    Exit Sub
End Sub

Private Sub RGBsplit(ByVal Col, r As Integer, g As Integer, b As Integer)

On Error GoTo errhandler

  b = (Col And 16711680) / 65536
  g = (Col And 65280) / 256
  r = Col And 255

errhandler:
    Exit Sub
End Sub

Private Sub GetComplementaryColor(ByVal BackGroundColor As Long, r As Byte, g As Byte, b As Byte)

On Error GoTo errhandler
    
    b = 255 - ((BackGroundColor And 16711680) / 65536)
    g = 255 - ((BackGroundColor And 65280) / 256)
    r = 255 - (BackGroundColor And 255)
       
errhandler:
    Exit Sub
End Sub

Public Sub color_Gradient()

On Error GoTo errhandler

Dim N As Long
Dim vRed            As Double
Dim vGreen          As Double
Dim vBlue           As Double
Dim kRed            As Double
Dim kGreen          As Double
Dim kBlue           As Double

Dim HscrollMax      As Integer
Dim HscrollMin      As Integer

Dim vHeight         As Long
    
    vHeight = 254
    ReDim C(0 To vHeight, 0 To 3)
    
    vRed = CDbl(VScroll1(0).Value)
    vGreen = CDbl(VScroll1(1).Value)
    vBlue = CDbl(VScroll1(2).Value)
        
    kRed = (255 - vRed) / vHeight
    kGreen = (255 - vGreen) / vHeight
    kBlue = (255 - vBlue) / vHeight
    
    ' set backcolor
    Picture1.BackColor = (vRed * 256 * 256) + (vGreen * 256) + vBlue
    
    ' calculate complementary color
    Picture4.Visible = False
    Picture4.BackColor = ((255 - vRed) * 256 * 256) + ((255 - vGreen) * 256) + 255 - vBlue
    Picture4.Visible = True
    
    HScroll1.Value = 0
    HScroll1_Change
    
    ' format color hex value
    txtPos(0).Text = " " & FormatHexCodeInv(CStr(Hex(Picture1.BackColor)))
    txt1DecB.Text = Format$(vRed, "000")
    txt1DecG.Text = Format$(vGreen, "000")
    txt1DecR.Text = Format$(vBlue, "000")
    
    txtPos(4).Text = " " & FormatHexCodeInv(CStr(Hex(Picture4.BackColor)))
    txt4DecB.Text = Format$(255 - vRed, "000")
    txt4DecG.Text = Format$(255 - vGreen, "000")
    txt4DecR.Text = Format$(255 - vBlue, "000")
    
    For N = 0 To vHeight Step 1
    
        Me.Line (213, N + 9)-(243, N + 9), RGB(CInt(vBlue), CInt(vGreen), CInt(vRed)), BF
        
        C(N, 0) = RGB(CInt(vBlue), CInt(vGreen), CInt(vRed))
        C(N, 1) = CInt(vBlue)
        C(N, 2) = CInt(vGreen)
        C(N, 3) = CInt(vRed)
        
        vRed = vRed + kRed: If vRed > 255 Then vRed = 255
        vGreen = vGreen + kGreen: If vGreen > 255 Then vGreen = 255
        vBlue = vBlue + kBlue: If vBlue > 255 Then vBlue = 255
        
    Next N
    
errhandler:
    Exit Sub
End Sub



Private Sub color_Implement(ByVal ColorClass As Long)
    
On Error GoTo errhandler

Dim N                   As Long
Dim P                   As Long
Dim tmp                 As String
Dim vRed                As Integer
Dim vGreen              As Integer
Dim vBlue               As Integer
Dim tRed                As Integer
Dim tGreen              As Integer
Dim tBlue               As Integer

Dim x()                 As String
Dim PosInRange          As Long

    ' set file_dirty flag
    File_Dirty_G = True
        
    ' reset adjustment settings
    frmMain.cmdReset_Click 1
    
    ' disable synchronising function
    frmMain.chkLock(0).Value = 0
    frmMain.chkLock(1).Value = 0
    frmMain.chkLock(2).Value = 0
    
    If ColorClass = 1 Then
        vRed = CInt(txt1DecR.Text)
        vGreen = CInt(txt1DecG.Text)
        vBlue = CInt(txt1DecB.Text)
    ElseIf ColorClass = 4 Then
        vRed = CInt(txt4DecR.Text)
        vGreen = CInt(txt4DecG.Text)
        vBlue = CInt(txt4DecB.Text)
    Else
        vRed = CInt(txt3DecR.Text)
        vGreen = CInt(txt3DecG.Text)
        vBlue = CInt(txt3DecB.Text)
    End If

    ' determines how many color fields to insert in-front of the selected Final color
    tRed = vRed
    tGreen = vGreen
    tBlue = vBlue
    PosInRange = 0
    For N = 1 To 18
        If tRed + int_Gradient_Interval < 256 And _
           tGreen + int_Gradient_Interval < 256 And _
           tBlue + int_Gradient_Interval < 256 And _
           PosInRange < 10 Then
            tRed = tRed + int_Gradient_Interval
            tGreen = tGreen + int_Gradient_Interval
            tBlue = tBlue + int_Gradient_Interval
            PosInRange = PosInRange + 1
        Else
            Exit For
        End If
    Next N
    
    ' adjust color range so that range starts at Index 0
    vRed = vRed + PosInRange * int_Gradient_Interval
    vGreen = vGreen + PosInRange * int_Gradient_Interval
    vBlue = vBlue + PosInRange * int_Gradient_Interval
        
    ' set scroll bars for selected Final color
    frmMain.ColorSelect(0).Value = vRed
    frmMain.ColorSelect(1).Value = vGreen
    frmMain.ColorSelect(2).Value = vBlue
    
    ' color coding active color field
    For P = 0 To 10
        frmMain.lblUsed(P).Visible = False
        frmMain.txtColorCode(P).BackColor = &HE9F3F4
        frmMain.lblRed(P).BackColor = &HE9F3F4
        frmMain.lblGreen(P).BackColor = &HE9F3F4
        frmMain.lblBlue(P).BackColor = &HE9F3F4
    Next P
    frmMain.lblUsed(PosInRange).Visible = True
    frmMain.txtColorCode(PosInRange).BackColor = &HFFFFFF
    frmMain.lblRed(PosInRange).BackColor = &HFFFFFF
    frmMain.lblGreen(PosInRange).BackColor = &HFFFFFF
    frmMain.lblBlue(PosInRange).BackColor = &HFFFFFF
        
errhandler:
    Exit Sub
End Sub


Private Sub TextBackColorWhite()

On Error GoTo errhandler

    txtPos(0).BackColor = &HFFFFFF
    txtPos(3).BackColor = &HFFFFFF
    txtPos(4).BackColor = &HFFFFFF
    
    txt1DecR.BackColor = &HFFFFFF
    txt1DecG.BackColor = &HFFFFFF
    txt1DecB.BackColor = &HFFFFFF
    
    txt3DecR.BackColor = &HFFFFFF
    txt3DecG.BackColor = &HFFFFFF
    txt3DecB.BackColor = &HFFFFFF

    txt4DecR.BackColor = &HFFFFFF
    txt4DecG.BackColor = &HFFFFFF
    txt4DecB.BackColor = &HFFFFFF
    
errhandler:
    Exit Sub
End Sub

Private Sub cmdAction_Click(Index As Integer)
    
On Error GoTo errhandler
        
    Select Case Index
        Case 1                                                      ' Show frmMain
            If cmdAction(1).Caption = "Open Editor" Then
                frmMain.WindowState = vbNormal
                frmMain.Top = 0
                frmMain.Left = 0
                frmMain.Visible = True
                cmdAction(1).Caption = "Close Editor"
                cmdAction(1).ToolTipText = " Close Main Color editor "
                
            ElseIf cmdAction(1).Caption = "Close Editor" Then
                frmMain.WindowState = vbNormal
                frmMain.Visible = False
                cmdAction(1).Caption = "Open Editor"
                cmdAction(1).ToolTipText = " Close Main Color editor "
            End If
            
        Case 2                                                      ' Minimise frmColor
            frmColor.WindowState = vbMinimized
            frmMain.WindowState = vbMinimized
            cmdAction(1).Caption = "Open Editor"
        
        Case 3                                                      ' Clear previous color
            Picture3.BackColor = Picture1.BackColor
            txtPos(3).Text = txtPos(0).Text
            txt3DecR.Text = txt1DecR.Text
            txt3DecG.Text = txt1DecG.Text
            txt3DecB.Text = txt1DecB.Text
            
        Case 4                                                      ' Store selected color
                        
        Case 5
            Unload frmMain
            
    End Select
    
errhandler:
    Exit Sub
End Sub

Private Sub cmdAction_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo errhandler

Dim txt1 As String
Dim txt2 As String
Dim txt3 As String

    txt1 = Trim$(txtPos(0).Text)
    txt2 = Left$(txt1, 2) & Space(1) & Mid$(txt1, 3, 2) & Space(1) & Right$(txt1, 2)
    txt3 = CInt("&H" & Left$(txt1, 2)) & ", " & CInt("&H" & Mid$(txt1, 3, 2)) & ", " & CInt("&H" & Right$(txt1, 2))
    
    Clipboard.Clear
    frmMain.lblClipBoard(0).Caption = vbNullString
    frmMain.lblClipBoard(1).Caption = vbNullString
    
    Clipboard.SetText txt1, 1
                              
    frmMain.lblClipBoard(0).Caption = txt2
    frmMain.lblClipBoard(1).Caption = txt3
    
    If Index = 4 Then
        With frmMain
            .Visible = True
            .WindowState = vbNormal
            .StatusBar1.style = 0
            .Width = 9825
            .ViewItem(0).Checked = True
            .ViewItem(1).Checked = False
        End With
        cmdAction(1).Caption = "Close Editor"
        LastActiveIndex_G = -1
        
        If Button = 2 Then
            color_Implement (4)
        ElseIf Shift = 1 Then
            color_Implement (3)
        Else
            color_Implement (1)
        End If
    End If
        
errhandler:
    Exit Sub
End Sub



Private Sub Form_Activate()

On Error GoTo errhandler

    color_Gradient
    LockGradient_M = False
        
errhandler:
    Exit Sub
End Sub

Private Sub Form_Load()

On Error GoTo errhandler

    form_StayOnTop frmColor, "ABSOLUTE", "TR"
        
    VScroll1(0).Value = CLng(frmMain.lblBlue(0).Caption)
    VScroll1(1).Value = CLng(frmMain.lblGreen(0).Caption)
    VScroll1(2).Value = CLng(frmMain.lblRed(0).Caption)
    
    Picture3.BackColor = Picture1.BackColor
    
    Call TextBackColorWhite
    
    txtPos(3).Text = txtPos(0).Text
    txt3DecR.Text = txt1DecR.Text
    txt3DecG.Text = txt1DecG.Text
    txt3DecB.Text = txt1DecB.Text
            
    Timer1.Interval = 50
    Timer1.Enabled = False
    
errhandler:
    Exit Sub
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo errhandler
    
    If LockGradient_M Then
        txtPos(2).Top = 2
        txtPos(2).Visible = True
        LockGradient_M = False
        Exit Sub
    End If
    
    If x < 214 Or x > 245 Or y < 9 Or y > 264 Then
        Me.Line (245, 1)-(252, 264), &HE9F3F4, BF
        txtPos(2).Visible = False
    Else
        If y > 255 + 8 Then y = 255 + 8
        
        txtPos(2).Text = " " & FormatHexCodeInv(Hex(C(y - 9, 0)))
        
        Me.Line (245, 9)-(252, 9), &HE9F3F4, BF
        Me.Line (245, y + 1)-(252, y + 20), &HE9F3F4, BF
        Me.Line (245, y)-(252, y), RGB(0, 0, 0), BF
        Me.Line (245, y - 1)-(252, y - 20), &HE9F3F4, BF
        
        txtPos(2).Top = y - 7
        txtPos(2).Visible = True
        
    End If

errhandler:
    Exit Sub
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo errhandler
            
Dim r                       As Integer
Dim g                       As Integer
Dim b                       As Integer

    If x < 214 Or x > 246 Or y < 9 Or y > 263 Then Exit Sub
    
    Picture1.BackColor = C(y - 9, 0)
    
    Call RGBSplit_Complement(Picture1.BackColor, r, g, b)
    
    Picture4.BackColor = RGB(r, g, b)
    
    VScroll1(2).Value = C(y - 9, 1)
    VScroll1(1).Value = C(y - 9, 2)
    VScroll1(0).Value = C(y - 9, 3)
        
    Me.Line (245, 1)-(252, 255), &HE9F3F4, BF
    Me.Line (245, 9)-(252, 9), RGB(0, 0, 0), BF
    
    txtPos(2).Top = 2
    LockGradient_M = True
    
errhandler:
    Exit Sub
End Sub


Private Sub HScroll1_Change()
    
On Error GoTo errhandler

Dim r As Integer
Dim g As Integer
Dim b As Integer
    
    Call RGBSplit_Complement(Picture1.BackColor, r, g, b)
    
    r = r + HScroll1.Value: If r > 255 Then r = 255
    g = g + HScroll1.Value: If g > 255 Then g = 255
    b = b + HScroll1.Value: If b > 255 Then b = 255

    Picture4.BackColor = RGB(r, g, b)
    txtPos(4).Text = " " & FormatHexCodeInv(CStr(Hex(Picture4.BackColor)))
    txt4DecR.Text = Format$(r, "000")
    txt4DecG.Text = Format$(g, "000")
    txt4DecB.Text = Format$(b, "000")
    
errhandler:
    Exit Sub
End Sub

Private Sub HScroll1_Scroll()
    
On Error GoTo errhandler

Dim r As Integer
Dim g As Integer
Dim b As Integer
    
    Call RGBSplit_Complement(Picture1.BackColor, r, g, b)
    
    r = r + HScroll1.Value: If r > 255 Then r = 255
    g = g + HScroll1.Value: If g > 255 Then g = 255
    b = b + HScroll1.Value: If b > 255 Then b = 255

    Picture4.BackColor = RGB(r, g, b)
    txtPos(4).Text = " " & FormatHexCodeInv(CStr(Hex(Picture4.BackColor)))
    txt4DecR.Text = Format$(r, "000")
    txt4DecG.Text = Format$(g, "000")
    txt4DecB.Text = Format$(b, "000")
    
errhandler:
    Exit Sub
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo errhandler
    
    If Shift Then
        Call PipetteOn(True)
    Else
        Call PipetteOn
    End If

errhandler:
    Exit Sub
End Sub


Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo errhandler

    Call PipetteOff

errhandler:
    Exit Sub
End Sub


Private Sub Timer1_Timer()

On Error GoTo errhandler

Dim hWndp               As Long
Dim hDCp                As Long
Dim Result              As Long
Dim Pt                  As POINTAPI

Dim r                   As Integer
Dim g                   As Integer
Dim b                   As Integer
  
Static LastX As Long
Static LastY As Long
        
    Call GetCursorPos(Pt)
    
    If Pt.x = LastX And Pt.y = LastY Then Exit Sub
      
    ' correct for icon size (field size = 32px, icon size = 17px)
    Pt.x = Pt.x - 17
    Pt.y = Pt.y + 15
          
    lblX.Caption = Pt.x
    lblY.Caption = Pt.y
    
    PosX_M = Pt.x
    PosY_M = Pt.y
          
    LastX = Pt.x
    LastY = Pt.y
        
    ' prevent flickering of color/complementary color fields
    If Pt.x > ConvertTwipsToPixels(Me.Left + 130, 0) And Pt.x < ConvertTwipsToPixels(Me.Left + 2200, 0) And _
       Pt.y > ConvertTwipsToPixels(Me.Top + 3600, 1) And Pt.y < ConvertTwipsToPixels(Me.Top + 4120, 1) Then Exit Sub
      
    hWndp = WindowFromPoint(Pt.x, Pt.y)
  
    hDCp = GetDC(hWndp)
  
    Call ScreenToClient(hWndp, Pt)
  
    Result = GetPixel(hDCp, Pt.x, Pt.y)
    If Result = -1 Then
        Call BitBlt(Picture1.hdc, 0, 0, 1, 1, hDCp, Pt.x, Pt.y, vbSrcCopy)
        Result = Picture1.Point(0, 0)
    End If
  
    Call ReleaseDC(hWndp, hDCp)
    
    If Result = -1 Then Exit Sub
    
    Picture1.BackColor = Result
                
    Call RGBsplit(Result, r, g, b)
        
    VScroll1(0).Value = b
    VScroll1(1).Value = g
    VScroll1(2).Value = r
    
    txtPos(0).Text = Space$(1) & CheckNull(Hex(r)) & CheckNull(Hex(g)) & CheckNull(Hex(b))
    txt1DecR.Text = Format$(r, "000")
    txt1DecG.Text = Format$(g, "000")
    txt1DecB.Text = Format$(b, "000")
    
    txtPos(4).Text = Space$(1) & CheckNull(Hex(255 - r)) & CheckNull(Hex(255 - g)) & CheckNull(Hex(255 - b))
    txt4DecR.Text = Format$(255 - r, "000")
    txt4DecG.Text = Format$(255 - g, "000")
    txt4DecB.Text = Format$(255 - b, "000")
    
    Picture4.BackColor = RGB(255 - r, 255 - g, 255 - b)
    
errhandler:
    Exit Sub
End Sub


Private Sub txt1DecB_Click()
    
    Clipboard.Clear
    Clipboard.SetText Trim$(txt1DecB.Text), 1
    Call TextBackColorWhite
    txt1DecB.BackColor = &HC0FFFF
    
End Sub

Private Sub txt1DecG_Click()
    
    Clipboard.Clear
    Clipboard.SetText Trim$(txt1DecG.Text), 1
    Call TextBackColorWhite
    txt1DecG.BackColor = &HC0FFFF
    
End Sub

Private Sub txt1DecR_Click()
    
    Clipboard.Clear
    Clipboard.SetText Trim$(txt1DecR.Text), 1
    Call TextBackColorWhite
    txt1DecR.BackColor = &HC0FFFF
    
End Sub

Private Sub txt3DecB_Click()
    
    Clipboard.Clear
    Clipboard.SetText Trim$(txt3DecB.Text), 1
    Call TextBackColorWhite
    txt3DecB.BackColor = &HEDE7D4
    
End Sub

Private Sub txt3DecG_Click()
    
    Clipboard.Clear
    Clipboard.SetText Trim$(txt3DecG.Text), 1
    Call TextBackColorWhite
    txt3DecG.BackColor = &HEDE7D4
    
End Sub

Private Sub txt3DecR_Click()
        
    Clipboard.Clear
    Clipboard.SetText Trim$(txt3DecR.Text), 1
    Call TextBackColorWhite
    txt3DecR.BackColor = &HEDE7D4
    
End Sub

Private Sub txt4DecB_Click()
    
    Clipboard.Clear
    Clipboard.SetText Trim$(txt4DecB.Text), 1
    Call TextBackColorWhite
    txt4DecB.BackColor = &HC9FFD2
    
End Sub


Private Sub txt4DecG_Click()
    
    Clipboard.Clear
    Clipboard.SetText Trim$(txt4DecG.Text), 1
    Call TextBackColorWhite
    txt4DecG.BackColor = &HC9FFD2
    
End Sub


Private Sub txt4DecR_Click()
    
    Clipboard.Clear
    Clipboard.SetText Trim$(txt4DecR.Text), 1
    Call TextBackColorWhite
    txt4DecR.BackColor = &HC9FFD2
    
End Sub

Private Sub txtPos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo errhandler

Dim txt1 As String
Dim txt2 As String
Dim txt3 As String
    
' SEL COLOR
    If Index = 0 Then
        txt1 = Trim$(txtPos(0).Text)    ' hex: FFFFFF
        txt2 = Left$(txt1, 2) & Space(1) & Mid$(txt1, 3, 2) & Space(1) & Right$(txt1, 2)    ' hex: FF FF FF
        txt3 = CInt("&H" & Left$(txt1, 2)) & ", " & CInt("&H" & Mid$(txt1, 3, 2)) & ", " & CInt("&H" & Right$(txt1, 2)) ' decimal: 255,255,255
        
        Clipboard.Clear
        frmMain.lblClipBoard(0).Caption = vbNullString
        frmMain.lblClipBoard(1).Caption = vbNullString
        
        Clipboard.SetText txt1, 1
                                  
        frmMain.lblClipBoard(0).Caption = txt2
        frmMain.lblClipBoard(1).Caption = txt3
        
        ' swop two first and two last characters when usign with Visual Basic
        If Button = 2 Then
            txt1 = Trim$(txtPos(0).Text)    ' hex: FFFFFF
            txt1 = Right$(txt1, 2) & Mid$(txt1, 3, 2) & Left$(txt1, 2)
            Clipboard.SetText txt1, 1
        End If
        
' COMPLEMENT
    ElseIf Index = 4 Then
        txt1 = Trim$(txtPos(4).Text)    ' hex: FFFFFF
        txt2 = Left$(txt1, 2) & Space(1) & Mid$(txt1, 3, 2) & Space(1) & Right$(txt1, 2)    ' hex: FF FF FF
        txt3 = CInt("&H" & Left$(txt1, 2)) & ", " & CInt("&H" & Mid$(txt1, 3, 2)) & ", " & CInt("&H" & Right$(txt1, 2)) ' decimal: 255,255,255
        
        Clipboard.Clear
        frmMain.lblClipBoard(0).Caption = vbNullString
        frmMain.lblClipBoard(1).Caption = vbNullString
        
        Clipboard.SetText txt1, 1
                                  
        frmMain.lblClipBoard(0).Caption = txt2
        frmMain.lblClipBoard(1).Caption = txt3
        
        ' swop two first and two last characters when usign with Visual Basic
        If Button = 2 Then
            txt1 = Trim$(txtPos(4).Text)    ' hex: FFFFFF
            txt1 = Right$(txt1, 2) & Mid$(txt1, 3, 2) & Left$(txt1, 2)
            Clipboard.SetText txt1, 1
        End If
        
' PREVIOUS
    ElseIf Index = 3 Then
        txt1 = Trim$(txtPos(3).Text)    ' hex: FFFFFF
        txt2 = Left$(txt1, 2) & Space(1) & Mid$(txt1, 3, 2) & Space(1) & Right$(txt1, 2)    ' hex: FF FF FF
        txt3 = CInt("&H" & Left$(txt1, 2)) & ", " & CInt("&H" & Mid$(txt1, 3, 2)) & ", " & CInt("&H" & Right$(txt1, 2)) ' decimal: 255,255,255
        
        Clipboard.Clear
        frmMain.lblClipBoard(0).Caption = vbNullString
        frmMain.lblClipBoard(1).Caption = vbNullString
        
        Clipboard.SetText txt1, 1
                                  
        frmMain.lblClipBoard(0).Caption = txt2
        frmMain.lblClipBoard(1).Caption = txt3
        
        ' swop two first and two last characters when usign with Visual Basic
        If Button = 2 Then
            txt1 = Trim$(txtPos(3).Text)    ' hex: FFFFFF
            txt1 = Right$(txt1, 2) & Mid$(txt1, 3, 2) & Left$(txt1, 2)
            Clipboard.SetText txt1, 1
        End If
        
    End If
   
    Call TextBackColorWhite
    Select Case Index
        Case 0: txtPos(Index).BackColor = &HC0FFFF
        Case 3: txtPos(Index).BackColor = &HEDE7D4
        Case 4: txtPos(Index).BackColor = &HC9FFD2
    End Select
    
errhandler:
    Exit Sub
End Sub


Private Sub VScroll1_Change(Index As Integer)

On Error GoTo errhandler

    color_Gradient

errhandler:
    Exit Sub
End Sub

Private Sub VScroll1_Scroll(Index As Integer)

On Error GoTo errhandler

    color_Gradient

errhandler:
    Exit Sub
End Sub

