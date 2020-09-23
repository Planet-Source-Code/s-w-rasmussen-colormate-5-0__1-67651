VERSION 5.00
Begin VB.Form frmLoad 
   Appearance      =   0  'Flat
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3495
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4485
   ControlBox      =   0   'False
   FillColor       =   &H000000C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   Icon            =   "Load.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Load.frx":08CA
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3840
      Picture         =   "Load.frx":346BE
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   2820
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   60
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " © 2008 by S.W. Rasmussen. All rights reserved"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   45
      TabIndex        =   0
      Top             =   3300
      Width           =   3720
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2

Private Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
  (ByVal hwnd As Long, _
   ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
  (ByVal hwnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
   
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal crKey As Long, _
   ByVal bAlpha As Long, _
   ByVal dwFlags As Long) As Long

Private Function AdjustWindowStyle()

Dim style As Long

  'in order to have transparent windows, the WS_EX_LAYERED window style must be applied to the form
   style = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
   
   If Not (style And WS_EX_LAYERED = WS_EX_LAYERED) Then
      style = style Or WS_EX_LAYERED
      SetWindowLong Me.hwnd, GWL_EXSTYLE, style
   End If
   
End Function

Public Sub SplashForm_FadeIn()

On Error GoTo errhandler
            
    'just adjust the window style and use a timer to fade the window in
     Call AdjustWindowStyle
     Timer1.Interval = 30
     Timer1.Enabled = True
       
errhandler:
    Exit Sub
       
End Sub
Private Sub Form_Load()

On Error GoTo errhandler

Dim start As Single

    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
            
    Label2.Caption = "version " & App.Major & "." & App.Minor & "." & App.Revision
            
    AdjustWindowStyle
    SplashForm_FadeIn

    Me.Show
    DoEvents
        
    start = Timer
    Do While Timer - start < 3
        DoEvents
    Loop
    
    ' clear file_dirty flag
    File_Dirty_G = False
        
    Me.Hide
    DoEvents
    
    frmColor.Show
        
    Unload Me
    
errhandler:
    Exit Sub
End Sub


Private Sub Timer1_Timer()

On Error GoTo errhandler

Dim alpha As Long
Static fadeValue As Long
   
    If (fadeValue + (256 * 0.05)) >= 256 Then
        Timer1.Enabled = False
        fadeValue = 0
        alpha = 255
    Else
          fadeValue = fadeValue + (256 * 0.05)
          alpha = fadeValue
    End If

    SetLayeredWindowAttributes Me.hwnd, 0&, alpha, LWA_ALPHA

errhandler:
    Exit Sub
    
End Sub


