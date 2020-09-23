VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " About ColorMate"
   ClientHeight    =   3690
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
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   3300
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "About.frx":08CA
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   120
      X2              =   3990
      Y1              =   3400
      Y2              =   3400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   120
      X2              =   120
      Y1              =   720
      Y2              =   3400
   End
   Begin VB.Label Label2 
      Caption         =   "© 2007 by S.W. Rasmussen. All rights reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   3480
      Width           =   3315
   End
   Begin VB.Label lblAbout 
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   2500
      Left            =   420
      TabIndex        =   1
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "ColorMate ver. 4.5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   780
      TabIndex        =   0
      Top             =   180
      Width           =   3555
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

On Error GoTo errhandler
    
    Unload Me
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Load()
    
    form_StayOnTop frmAbout, "ABSOLUTE", "C"
    
    Label1.Caption = "ColorMate ver. " & App.Major & "." & App.Minor & "." & App.Revision
    
    lblAbout.Caption = "This application is intended to assist you managing " & _
                       "color usage in your web design." & vbCrLf & vbCrLf & _
                       "Grab color information from any pixel on your screen. Store the color in an " & _
                       "advanced editor for flexible adjustment of colors." & vbCrLf & vbCrLf & _
                       "Up to 10 colors can be saved and later loaded into ColorMate facilitating " & _
                       "consistent management of colors on your web pages." & vbCrLf & vbCrLf & _
                       "Suggestions and idears to:" & vbCrLf & _
                       "Søren W. Rasmussen - swr@seqtools.dk"

    
End Sub

