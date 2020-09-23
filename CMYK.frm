VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCMYK 
   BackColor       =   &H00878724&
   Caption         =   "Convert to CMYK"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3615
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   300
      Left            =   4320
      TabIndex        =   15
      Top             =   3180
      Width           =   855
   End
   Begin RichTextLib.RichTextBox txtColorValue 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393217
      TextRTF         =   $"CMYK.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "Blue"
      Height          =   195
      Index           =   6
      Left            =   2160
      TabIndex        =   14
      Top             =   2460
      Width           =   660
   End
   Begin VB.Label lblRGB 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   3180
      TabIndex        =   13
      Top             =   2460
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Green"
      Height          =   195
      Index           =   5
      Left            =   2160
      TabIndex        =   12
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label lblRGB 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   3180
      TabIndex        =   11
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   195
      Index           =   4
      Left            =   2160
      TabIndex        =   10
      Top             =   1860
      Width           =   660
   End
   Begin VB.Label lblRGB 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3180
      TabIndex        =   9
      Top             =   1860
      Width           =   600
   End
   Begin VB.Label lblCMYKvalue 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   3180
      TabIndex        =   7
      Top             =   960
      Width           =   600
   End
   Begin VB.Label lblCMYKvalue 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   3180
      TabIndex        =   6
      Top             =   600
      Width           =   600
   End
   Begin VB.Label lblCMYKvalue 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3180
      TabIndex        =   5
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Magenta"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Yellow"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Black"
      Height          =   195
      Index           =   3
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label lblCMYKvalue 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   3180
      TabIndex        =   1
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Cyan"
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   660
   End
End
Attribute VB_Name = "frmCMYK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()
    
Dim ColorValue As Long
    
    If Trim$(txtColorValue.Text) = vbNullString Then Exit Sub
    
    ColorValue = CLng(txtColorValue.Text)
    
    lblCMYKvalue(0).Caption = jwlCyan(ColorValue)
    lblCMYKvalue(1).Caption = jwlMagenta(ColorValue)
    lblCMYKvalue(2).Caption = jwlYellow(ColorValue)
    lblCMYKvalue(3).Caption = jwlBlack(ColorValue)
    
    lblRGB(0).Caption = jwlRed(ColorValue)
    lblRGB(1).Caption = jwlGreen(ColorValue)
    lblRGB(2).Caption = jwlBlue(ColorValue)
    
End Sub


