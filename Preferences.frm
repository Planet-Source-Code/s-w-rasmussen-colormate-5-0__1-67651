VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPreferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Preferences"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Preferences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Top             =   2940
      Width           =   795
   End
   Begin RichTextLib.RichTextBox txtPref 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"Preferences.frx":08CA
   End
   Begin RichTextLib.RichTextBox txtPref 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"Preferences.frx":0945
   End
   Begin RichTextLib.RichTextBox txtPref 
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Preferences.frx":09C0
   End
   Begin VB.Label Label1 
      Caption         =   "User name:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Initials:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   2715
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()

On Error GoTo errhandler

    Unload Me
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Load()
    
On Error GoTo errhandler
    
    form_StayOnTop frmPreferences, "ABSOLUTE", "C"
        
    txtPref(0).Text = pref_UserName
    txtPref(1).Text = pref_Description
    txtPref(2).Text = pref_Initials
    
errhandler:
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    pref_UserName = txtPref(0).Text
    pref_Description = txtPref(1).Text
    pref_Initials = txtPref(2).Text
    
End Sub


Private Sub txtPref_KeyPress(Index As Integer, KeyAscii As Integer)
    
On Error GoTo errhandler

    If Index = 2 Then
        ' convert to upper case
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    End If
    
    ' set file_dirty flag
    File_Dirty_G = True

errhandler:
    Exit Sub
End Sub

