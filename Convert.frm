VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Convert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Color Codes"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Convert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txt0DecR 
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Convert.frx":000C
   End
   Begin RichTextLib.RichTextBox txtHex0 
      Height          =   255
      Left            =   1620
      TabIndex        =   2
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"Convert.frx":0087
   End
   Begin RichTextLib.RichTextBox txtHex3 
      Height          =   255
      Left            =   1620
      TabIndex        =   3
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"Convert.frx":0102
   End
   Begin RichTextLib.RichTextBox txt0DecG 
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Convert.frx":017D
   End
   Begin RichTextLib.RichTextBox txt0DecB 
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Convert.frx":01F8
   End
   Begin RichTextLib.RichTextBox txt3DecR 
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   480
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Convert.frx":0273
   End
   Begin RichTextLib.RichTextBox txt3DecG 
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   480
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Convert.frx":02EE
   End
   Begin RichTextLib.RichTextBox txt3DecB 
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   480
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   450
      _Version        =   393217
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   4
      Appearance      =   0
      TextRTF         =   $"Convert.frx":0369
   End
   Begin VB.Label Label1 
      Caption         =   "Complementary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   1500
   End
   Begin VB.Label txtDec3B 
      Caption         =   "Selected Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1500
   End
End
Attribute VB_Name = "Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    
    frmColor.color_Gradient
    
End Sub

