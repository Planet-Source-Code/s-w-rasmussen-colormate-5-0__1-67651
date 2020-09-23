VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E9F3F4&
   BorderStyle     =   0  'None
   Caption         =   " ColorMate"
   ClientHeight    =   7545
   ClientLeft      =   2370
   ClientTop       =   720
   ClientWidth     =   9735
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
   ForeColor       =   &H00000000&
   Icon            =   "ColorSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7110
      Left            =   5250
      TabIndex        =   117
      Top             =   45
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   12541
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   15332340
      TabCaption(0)   =   "Stored Colors"
      TabPicture(0)   =   "ColorSelect.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFinal(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFinal(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFinal(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblFinal(9)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFinal(10)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblFinal(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblFinal(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblFinal(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblFinal(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblFinal(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblTarget(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblTarget(6)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblTarget(7)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblTarget(8)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblTarget(9)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblFinalNo(10)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblFinalNo(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblFinalNo(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblFinalNo(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblFinalNo(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblFinalNo(5)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblFinalNo(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblFinalNo(3)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblFinalNo(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblFinalNo(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblTarget(4)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblTarget(2)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblTarget(3)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblTarget(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblTarget(5)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtFinal(20)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtFinal(19)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtFinal(18)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtFinal(17)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtFinal(16)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtFinal(15)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtFinal(14)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtFinal(13)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtFinal(12)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtFinal(11)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtFinal(10)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtFinal(9)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtFinal(8)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtFinal(7)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtFinal(6)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtFinal(5)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtFinal(4)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtFinal(3)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtFinal(2)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtFinal(1)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).ControlCount=   50
      TabCaption(1)   =   "Link"
      TabPicture(1)   =   "ColorSelect.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLinkColor(1)"
      Tab(1).Control(1)=   "lblLinkColor(2)"
      Tab(1).Control(2)=   "lblLinkColor(4)"
      Tab(1).Control(3)=   "lblLinkColor(3)"
      Tab(1).Control(4)=   "lblLinkColor(6)"
      Tab(1).Control(5)=   "lblLinkColor(5)"
      Tab(1).Control(6)=   "lblLinkColor(8)"
      Tab(1).Control(7)=   "lblLinkColor(7)"
      Tab(1).Control(8)=   "lblLinkColor(10)"
      Tab(1).Control(9)=   "lblLinkColor(9)"
      Tab(1).Control(10)=   "lblFinalNo(0)"
      Tab(1).Control(11)=   "lblFinalNo(11)"
      Tab(1).Control(12)=   "lblFinalNo(12)"
      Tab(1).Control(13)=   "lblFinalNo(13)"
      Tab(1).Control(14)=   "lblFinalNo(14)"
      Tab(1).Control(15)=   "txtLinkBackColor(10)"
      Tab(1).Control(16)=   "txtLinkForeColor(10)"
      Tab(1).Control(17)=   "txtLinkBackColor(9)"
      Tab(1).Control(18)=   "txtLinkForeColor(9)"
      Tab(1).Control(19)=   "txtLinkBackColor(8)"
      Tab(1).Control(20)=   "txtLinkForeColor(8)"
      Tab(1).Control(21)=   "txtLinkBackColor(7)"
      Tab(1).Control(22)=   "txtLinkForeColor(7)"
      Tab(1).Control(23)=   "txtLinkBackColor(6)"
      Tab(1).Control(24)=   "txtLinkForeColor(6)"
      Tab(1).Control(25)=   "txtLinkBackColor(5)"
      Tab(1).Control(26)=   "txtLinkForeColor(5)"
      Tab(1).Control(27)=   "txtLinkBackColor(4)"
      Tab(1).Control(28)=   "txtLinkForeColor(4)"
      Tab(1).Control(29)=   "txtLinkBackColor(3)"
      Tab(1).Control(30)=   "txtLinkForeColor(3)"
      Tab(1).Control(31)=   "txtLinkBackColor(2)"
      Tab(1).Control(32)=   "txtLinkForeColor(2)"
      Tab(1).Control(33)=   "txtLinkBackColor(1)"
      Tab(1).Control(34)=   "txtLinkBackColorHex(10)"
      Tab(1).Control(35)=   "txtLinkForeColorHex(10)"
      Tab(1).Control(36)=   "txtLinkBackColorHex(9)"
      Tab(1).Control(37)=   "txtLinkForeColorHex(9)"
      Tab(1).Control(38)=   "txtLinkBackColorHex(8)"
      Tab(1).Control(39)=   "txtLinkForeColorHex(8)"
      Tab(1).Control(40)=   "txtLinkBackColorHex(7)"
      Tab(1).Control(41)=   "txtLinkForeColorHex(7)"
      Tab(1).Control(42)=   "txtLinkBackColorHex(6)"
      Tab(1).Control(43)=   "txtLinkForeColorHex(6)"
      Tab(1).Control(44)=   "txtLinkBackColorHex(5)"
      Tab(1).Control(45)=   "txtLinkForeColorHex(5)"
      Tab(1).Control(46)=   "txtLinkBackColorHex(4)"
      Tab(1).Control(47)=   "txtLinkForeColorHex(4)"
      Tab(1).Control(48)=   "txtLinkBackColorHex(3)"
      Tab(1).Control(49)=   "txtLinkForeColorHex(3)"
      Tab(1).Control(50)=   "txtLinkBackColorHex(2)"
      Tab(1).Control(51)=   "txtLinkForeColorHex(2)"
      Tab(1).Control(52)=   "txtLinkBackColorHex(1)"
      Tab(1).Control(53)=   "txtLinkForeColorHex(1)"
      Tab(1).Control(54)=   "txtLinkForeColor(1)"
      Tab(1).ControlCount=   55
      TabCaption(2)   =   "Text"
      TabPicture(2)   =   "ColorSelect.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblTextColor(1)"
      Tab(2).Control(1)=   "lblTextColor(2)"
      Tab(2).Control(2)=   "lblTextColor(3)"
      Tab(2).Control(3)=   "lblTextColor(4)"
      Tab(2).Control(4)=   "lblTextColor(5)"
      Tab(2).Control(5)=   "lblFinalNo(15)"
      Tab(2).Control(6)=   "lblFinalNo(16)"
      Tab(2).Control(7)=   "lblFinalNo(17)"
      Tab(2).Control(8)=   "lblFinalNo(18)"
      Tab(2).Control(9)=   "lblFinalNo(19)"
      Tab(2).Control(10)=   "txtForeColor(5)"
      Tab(2).Control(11)=   "txtForeColorHex(5)"
      Tab(2).Control(12)=   "txtForeColor(4)"
      Tab(2).Control(13)=   "txtForeColorHex(4)"
      Tab(2).Control(14)=   "txtForeColor(3)"
      Tab(2).Control(15)=   "txtForeColorHex(3)"
      Tab(2).Control(16)=   "txtForeColor(2)"
      Tab(2).Control(17)=   "txtForeColorHex(2)"
      Tab(2).Control(18)=   "txtForeColor(1)"
      Tab(2).Control(19)=   "txtForeColorHex(1)"
      Tab(2).Control(20)=   "txtBackColor(5)"
      Tab(2).Control(21)=   "txtBackColor(4)"
      Tab(2).Control(22)=   "txtBackColor(3)"
      Tab(2).Control(23)=   "txtBackColor(2)"
      Tab(2).Control(24)=   "txtBackColor(1)"
      Tab(2).Control(25)=   "txtBackColorHex(5)"
      Tab(2).Control(26)=   "txtBackColorHex(4)"
      Tab(2).Control(27)=   "txtBackColorHex(3)"
      Tab(2).Control(28)=   "txtBackColorHex(2)"
      Tab(2).Control(29)=   "txtBackColorHex(1)"
      Tab(2).ControlCount=   30
      TabCaption(3)   =   "Web Elements"
      TabPicture(3)   =   "ColorSelect.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label"
      Tab(3).Control(1)=   "lstWebElements(0)"
      Tab(3).Control(2)=   "lstWebElements(1)"
      Tab(3).ControlCount=   3
      Begin VB.ListBox lstWebElements 
         Height          =   2010
         Index           =   1
         Left            =   -74820
         Sorted          =   -1  'True
         TabIndex        =   265
         Top             =   2280
         Width           =   4035
      End
      Begin VB.ListBox lstWebElements 
         Height          =   2010
         Index           =   0
         Left            =   -74820
         Sorted          =   -1  'True
         TabIndex        =   263
         Top             =   180
         Width           =   4035
      End
      Begin RichTextLib.RichTextBox txtLinkForeColor 
         Height          =   225
         Index           =   1
         Left            =   -73980
         TabIndex        =   198
         Tag             =   "1Normal_F"
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":093A
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   118
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   1110
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":09B5
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   119
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   2432
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0A30
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   120
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   3754
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0AAB
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   121
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   5076
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0B26
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   122
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   6400
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0BA1
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   6
         Left            =   2220
         TabIndex        =   123
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   1110
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0C1C
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   7
         Left            =   2220
         TabIndex        =   124
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   2432
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0C97
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   8
         Left            =   2220
         TabIndex        =   125
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   3754
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0D12
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   9
         Left            =   2220
         TabIndex        =   126
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   5076
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0D8D
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   10
         Left            =   2220
         TabIndex        =   127
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   6400
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0E08
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   11
         Left            =   1020
         TabIndex        =   128
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   1110
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0E83
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   12
         Left            =   1020
         TabIndex        =   129
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   2432
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0EFE
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   13
         Left            =   1020
         TabIndex        =   130
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   3754
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0F79
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   14
         Left            =   1020
         TabIndex        =   131
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   5076
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":0FF4
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   15
         Left            =   1020
         TabIndex        =   132
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   6400
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":106F
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   16
         Left            =   3120
         TabIndex        =   133
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   1110
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":10EA
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   17
         Left            =   3120
         TabIndex        =   134
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   2432
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1165
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   18
         Left            =   3120
         TabIndex        =   135
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   3754
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":11E0
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   19
         Left            =   3120
         TabIndex        =   136
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   5076
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":125B
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
      Begin RichTextLib.RichTextBox txtFinal 
         Height          =   225
         Index           =   20
         Left            =   3120
         TabIndex        =   137
         TabStop         =   0   'False
         ToolTipText     =   " Double-click to highlight, CTRL+C to copy "
         Top             =   6400
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":12D6
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
      Begin RichTextLib.RichTextBox txtLinkForeColorHex 
         Height          =   225
         Index           =   1
         Left            =   -74880
         TabIndex        =   160
         TabStop         =   0   'False
         Tag             =   "1Normal_F"
         Top             =   120
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1351
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
      Begin RichTextLib.RichTextBox txtLinkBackColorHex 
         Height          =   225
         Index           =   1
         Left            =   -74880
         TabIndex        =   161
         TabStop         =   0   'False
         Tag             =   "1B"
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":13CC
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
      Begin RichTextLib.RichTextBox txtLinkForeColorHex 
         Height          =   225
         Index           =   2
         Left            =   -74880
         TabIndex        =   162
         TabStop         =   0   'False
         Tag             =   "3a"
         Top             =   660
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1447
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
      Begin RichTextLib.RichTextBox txtLinkBackColorHex 
         Height          =   225
         Index           =   2
         Left            =   -74880
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   900
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":14C2
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
      Begin RichTextLib.RichTextBox txtLinkForeColorHex 
         Height          =   225
         Index           =   3
         Left            =   -74880
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   1380
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":153D
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
      Begin RichTextLib.RichTextBox txtLinkBackColorHex 
         Height          =   225
         Index           =   3
         Left            =   -74880
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   1620
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":15B8
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
      Begin RichTextLib.RichTextBox txtLinkForeColorHex 
         Height          =   225
         Index           =   4
         Left            =   -74880
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   1920
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1633
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
      Begin RichTextLib.RichTextBox txtLinkBackColorHex 
         Height          =   225
         Index           =   4
         Left            =   -74880
         TabIndex        =   167
         TabStop         =   0   'False
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":16AE
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
      Begin RichTextLib.RichTextBox txtLinkForeColorHex 
         Height          =   225
         Index           =   5
         Left            =   -74880
         TabIndex        =   170
         TabStop         =   0   'False
         Top             =   2640
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1729
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
      Begin RichTextLib.RichTextBox txtLinkBackColorHex 
         Height          =   225
         Index           =   5
         Left            =   -74880
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   2880
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":17A4
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
      Begin RichTextLib.RichTextBox txtLinkForeColorHex 
         Height          =   225
         Index           =   6
         Left            =   -74880
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   3180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":181F
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
      Begin RichTextLib.RichTextBox txtLinkBackColorHex 
         Height          =   225
         Index           =   6
         Left            =   -74880
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   3420
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":189A
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
      Begin RichTextLib.RichTextBox txtLinkForeColorHex 
         Height          =   225
         Index           =   7
         Left            =   -74880
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   3900
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1915
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
      Begin RichTextLib.RichTextBox txtLinkBackColorHex 
         Height          =   225
         Index           =   7
         Left            =   -74880
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   4140
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1990
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
      Begin RichTextLib.RichTextBox txtLinkForeColorHex 
         Height          =   225
         Index           =   8
         Left            =   -74880
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   4440
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1A0B
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
      Begin RichTextLib.RichTextBox txtLinkBackColorHex 
         Height          =   225
         Index           =   8
         Left            =   -74880
         TabIndex        =   179
         TabStop         =   0   'False
         Top             =   4680
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1A86
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
      Begin RichTextLib.RichTextBox txtLinkForeColorHex 
         Height          =   225
         Index           =   9
         Left            =   -74880
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   5160
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1B01
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
      Begin RichTextLib.RichTextBox txtLinkBackColorHex 
         Height          =   225
         Index           =   9
         Left            =   -74880
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   5400
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1B7C
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
      Begin RichTextLib.RichTextBox txtLinkForeColorHex 
         Height          =   225
         Index           =   10
         Left            =   -74880
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   5700
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1BF7
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
      Begin RichTextLib.RichTextBox txtLinkBackColorHex 
         Height          =   225
         Index           =   10
         Left            =   -74880
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   5940
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1C72
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtBackColorHex 
         Height          =   225
         Index           =   1
         Left            =   -74880
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1CED
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
      Begin RichTextLib.RichTextBox txtBackColorHex 
         Height          =   225
         Index           =   2
         Left            =   -74880
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   1620
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1D68
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
      Begin RichTextLib.RichTextBox txtBackColorHex 
         Height          =   225
         Index           =   3
         Left            =   -74880
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   2880
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1DE3
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
      Begin RichTextLib.RichTextBox txtBackColorHex 
         Height          =   225
         Index           =   4
         Left            =   -74880
         TabIndex        =   194
         TabStop         =   0   'False
         Top             =   4140
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1E5E
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
      Begin RichTextLib.RichTextBox txtBackColorHex 
         Height          =   225
         Index           =   5
         Left            =   -74880
         TabIndex        =   196
         TabStop         =   0   'False
         Top             =   5400
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1ED9
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
      Begin RichTextLib.RichTextBox txtLinkBackColor 
         Height          =   225
         Index           =   1
         Left            =   -73980
         TabIndex        =   199
         Tag             =   "1B"
         Top             =   360
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1F54
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
      Begin RichTextLib.RichTextBox txtLinkForeColor 
         Height          =   225
         Index           =   2
         Left            =   -73980
         TabIndex        =   200
         Tag             =   "3a"
         Top             =   660
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":1FCF
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
      Begin RichTextLib.RichTextBox txtLinkBackColor 
         Height          =   225
         Index           =   2
         Left            =   -73980
         TabIndex        =   201
         Top             =   900
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":204A
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
      Begin RichTextLib.RichTextBox txtLinkForeColor 
         Height          =   225
         Index           =   3
         Left            =   -73980
         TabIndex        =   202
         Top             =   1380
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":20C5
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
      Begin RichTextLib.RichTextBox txtLinkBackColor 
         Height          =   225
         Index           =   3
         Left            =   -73980
         TabIndex        =   203
         Top             =   1620
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2140
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
      Begin RichTextLib.RichTextBox txtLinkForeColor 
         Height          =   225
         Index           =   4
         Left            =   -73980
         TabIndex        =   204
         Top             =   1920
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":21BB
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
      Begin RichTextLib.RichTextBox txtLinkBackColor 
         Height          =   225
         Index           =   4
         Left            =   -73980
         TabIndex        =   205
         Top             =   2160
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2236
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
      Begin RichTextLib.RichTextBox txtLinkForeColor 
         Height          =   225
         Index           =   5
         Left            =   -73980
         TabIndex        =   206
         Top             =   2640
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":22B1
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
      Begin RichTextLib.RichTextBox txtLinkBackColor 
         Height          =   225
         Index           =   5
         Left            =   -73980
         TabIndex        =   207
         Top             =   2880
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":232C
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
      Begin RichTextLib.RichTextBox txtLinkForeColor 
         Height          =   225
         Index           =   6
         Left            =   -73980
         TabIndex        =   208
         Top             =   3180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":23A7
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
      Begin RichTextLib.RichTextBox txtLinkBackColor 
         Height          =   225
         Index           =   6
         Left            =   -73980
         TabIndex        =   209
         Top             =   3420
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2422
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
      Begin RichTextLib.RichTextBox txtLinkForeColor 
         Height          =   225
         Index           =   7
         Left            =   -73980
         TabIndex        =   210
         Top             =   3900
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":249D
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
      Begin RichTextLib.RichTextBox txtLinkBackColor 
         Height          =   225
         Index           =   7
         Left            =   -73980
         TabIndex        =   211
         Top             =   4140
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2518
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
      Begin RichTextLib.RichTextBox txtLinkForeColor 
         Height          =   225
         Index           =   8
         Left            =   -73980
         TabIndex        =   212
         Top             =   4440
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2593
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
      Begin RichTextLib.RichTextBox txtLinkBackColor 
         Height          =   225
         Index           =   8
         Left            =   -73980
         TabIndex        =   213
         Top             =   4680
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":260E
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
      Begin RichTextLib.RichTextBox txtLinkForeColor 
         Height          =   225
         Index           =   9
         Left            =   -73980
         TabIndex        =   214
         Top             =   5160
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2689
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
      Begin RichTextLib.RichTextBox txtLinkBackColor 
         Height          =   225
         Index           =   9
         Left            =   -73980
         TabIndex        =   215
         Top             =   5400
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2704
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
      Begin RichTextLib.RichTextBox txtLinkForeColor 
         Height          =   225
         Index           =   10
         Left            =   -73980
         TabIndex        =   216
         Top             =   5700
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":277F
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
      Begin RichTextLib.RichTextBox txtLinkBackColor 
         Height          =   225
         Index           =   10
         Left            =   -73980
         TabIndex        =   217
         Top             =   5940
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":27FA
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
      Begin RichTextLib.RichTextBox txtBackColor 
         Height          =   225
         Index           =   1
         Left            =   -73980
         TabIndex        =   218
         Top             =   360
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2875
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
      Begin RichTextLib.RichTextBox txtBackColor 
         Height          =   225
         Index           =   2
         Left            =   -73980
         TabIndex        =   219
         Top             =   1620
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":28F0
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
      Begin RichTextLib.RichTextBox txtBackColor 
         Height          =   225
         Index           =   3
         Left            =   -73980
         TabIndex        =   220
         Top             =   2880
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":296B
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
      Begin RichTextLib.RichTextBox txtBackColor 
         Height          =   225
         Index           =   4
         Left            =   -73980
         TabIndex        =   221
         Top             =   4140
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":29E6
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
      Begin RichTextLib.RichTextBox txtBackColor 
         Height          =   225
         Index           =   5
         Left            =   -73980
         TabIndex        =   222
         Top             =   5400
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   15527148
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2A61
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
      Begin RichTextLib.RichTextBox txtForeColorHex 
         Height          =   225
         Index           =   1
         Left            =   -74880
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   120
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2ADC
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
      Begin RichTextLib.RichTextBox txtForeColor 
         Height          =   225
         Index           =   1
         Left            =   -73980
         TabIndex        =   224
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2B57
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
      Begin RichTextLib.RichTextBox txtForeColorHex 
         Height          =   225
         Index           =   2
         Left            =   -74880
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   1380
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2BD2
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
      Begin RichTextLib.RichTextBox txtForeColor 
         Height          =   225
         Index           =   2
         Left            =   -73980
         TabIndex        =   226
         Top             =   1380
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2C4D
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
      Begin RichTextLib.RichTextBox txtForeColorHex 
         Height          =   225
         Index           =   3
         Left            =   -74880
         TabIndex        =   227
         TabStop         =   0   'False
         Top             =   2640
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2CC8
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
      Begin RichTextLib.RichTextBox txtForeColor 
         Height          =   225
         Index           =   3
         Left            =   -73980
         TabIndex        =   228
         Top             =   2640
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2D43
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
      Begin RichTextLib.RichTextBox txtForeColorHex 
         Height          =   225
         Index           =   4
         Left            =   -74880
         TabIndex        =   229
         TabStop         =   0   'False
         Top             =   3900
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2DBE
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
      Begin RichTextLib.RichTextBox txtForeColor 
         Height          =   225
         Index           =   4
         Left            =   -73980
         TabIndex        =   230
         Top             =   3900
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2E39
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
      Begin RichTextLib.RichTextBox txtForeColorHex 
         Height          =   225
         Index           =   5
         Left            =   -74880
         TabIndex        =   231
         TabStop         =   0   'False
         Top             =   5160
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         MaxLength       =   7
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2EB4
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
      Begin RichTextLib.RichTextBox txtForeColor 
         Height          =   225
         Index           =   5
         Left            =   -73980
         TabIndex        =   232
         Top             =   5160
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393217
         BackColor       =   16512495
         Appearance      =   0
         TextRTF         =   $"ColorSelect.frx":2F2F
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
      Begin VB.Label Label 
         Height          =   2295
         Left            =   -74820
         TabIndex        =   264
         Top             =   4380
         Width           =   4035
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   1500
         TabIndex        =   261
         Top             =   5400
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   1500
         TabIndex        =   260
         Top             =   120
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   1500
         TabIndex        =   259
         Top             =   2760
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   1500
         TabIndex        =   258
         Top             =   1440
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   1500
         TabIndex        =   257
         Top             =   4080
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   256
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   255
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   254
         Top             =   2760
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   253
         Top             =   4080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   252
         Top             =   5400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   19
         Left            =   -70995
         TabIndex        =   242
         Top             =   5160
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   -70995
         TabIndex        =   241
         Top             =   3900
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   -70995
         TabIndex        =   240
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   -70990
         TabIndex        =   239
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   -70995
         TabIndex        =   238
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   -70995
         TabIndex        =   237
         Top             =   5160
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   -70995
         TabIndex        =   236
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   -70995
         TabIndex        =   235
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   -70990
         TabIndex        =   234
         Top             =   3900
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   -70990
         TabIndex        =   233
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblTextColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This is a sample text   for evaluation of fore- and backgrund colors of your html pages. The color codes are"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1065
         Index           =   5
         Left            =   -72780
         TabIndex        =   197
         Top             =   5160
         Width           =   2040
      End
      Begin VB.Label lblTextColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This is a sample text   for evaluation of fore- and backgrund colors of your html pages. The color codes are"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1065
         Index           =   4
         Left            =   -72780
         TabIndex        =   195
         Top             =   3900
         Width           =   2040
      End
      Begin VB.Label lblTextColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This is a sample text   for evaluation of fore- and backgrund colors of your html pages. The color codes are"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1065
         Index           =   3
         Left            =   -72780
         TabIndex        =   193
         Top             =   2640
         Width           =   2040
      End
      Begin VB.Label lblTextColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This is a sample text   for evaluation of fore- and backgrund colors of your html pages. The color codes are"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1065
         Index           =   2
         Left            =   -72780
         TabIndex        =   191
         Top             =   1380
         Width           =   2040
      End
      Begin VB.Label lblLinkColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Normal ab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   470
         Index           =   9
         Left            =   -72780
         TabIndex        =   187
         Top             =   5160
         Width           =   2040
      End
      Begin VB.Label lblLinkColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hover abc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   470
         Index           =   10
         Left            =   -72780
         TabIndex        =   186
         Top             =   5700
         Width           =   2040
      End
      Begin VB.Label lblLinkColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Normal ab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   470
         Index           =   7
         Left            =   -72780
         TabIndex        =   181
         Top             =   3900
         Width           =   2040
      End
      Begin VB.Label lblLinkColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hover abc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   470
         Index           =   8
         Left            =   -72780
         TabIndex        =   180
         Top             =   4440
         Width           =   2040
      End
      Begin VB.Label lblLinkColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Normal ab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   470
         Index           =   5
         Left            =   -72780
         TabIndex        =   175
         Top             =   2640
         Width           =   2040
      End
      Begin VB.Label lblLinkColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hover abc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   470
         Index           =   6
         Left            =   -72780
         TabIndex        =   174
         Top             =   3180
         Width           =   2040
      End
      Begin VB.Label lblLinkColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Normal ab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   470
         Index           =   3
         Left            =   -72780
         TabIndex        =   169
         Top             =   1380
         Width           =   2040
      End
      Begin VB.Label lblLinkColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hover abc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   470
         Index           =   4
         Left            =   -72780
         TabIndex        =   168
         Top             =   1920
         Width           =   2040
      End
      Begin VB.Label lblLinkColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hover abc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   470
         Index           =   2
         Left            =   -72780
         TabIndex        =   159
         Top             =   660
         Width           =   2040
      End
      Begin VB.Label lblLinkColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Normal ab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   470
         Index           =   1
         Left            =   -72780
         TabIndex        =   158
         Top             =   120
         Width           =   2040
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   2220
         TabIndex        =   144
         Top             =   4080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   2220
         TabIndex        =   145
         Top             =   2760
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2220
         TabIndex        =   146
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   2220
         TabIndex        =   147
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblFinalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   2220
         TabIndex        =   143
         Top             =   5400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   3600
         TabIndex        =   139
         Top             =   4080
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   3600
         TabIndex        =   140
         Top             =   2760
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   3600
         TabIndex        =   141
         Top             =   1440
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   3600
         TabIndex        =   142
         Top             =   120
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   3600
         TabIndex        =   138
         Top             =   5400
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   1
         Left            =   120
         TabIndex        =   157
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   2
         Left            =   120
         TabIndex        =   156
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   3
         Left            =   120
         TabIndex        =   155
         Top             =   2760
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   4
         Left            =   120
         TabIndex        =   154
         Top             =   4080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   5
         Left            =   120
         TabIndex        =   153
         Top             =   5400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   10
         Left            =   2220
         TabIndex        =   152
         Top             =   5400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   9
         Left            =   2220
         TabIndex        =   151
         Top             =   4080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   8
         Left            =   2220
         TabIndex        =   150
         Top             =   2760
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   7
         Left            =   2220
         TabIndex        =   149
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   6
         Left            =   2220
         TabIndex        =   148
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblTextColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This is a sample text   for evaluation of fore- and backgrund colors of your html pages. The color codes are"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1065
         Index           =   1
         Left            =   -72780
         TabIndex        =   189
         Top             =   120
         Width           =   2040
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   116
      Top             =   7245
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10927
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtNewDec 
      Height          =   255
      Index           =   0
      Left            =   3060
      TabIndex        =   1
      ToolTipText     =   " Enter decimal value for red "
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      MaxLength       =   3
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":2FAA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   540
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtNewHexValue 
      Height          =   255
      Left            =   2340
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   " Enter RGB value (Hex) "
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   12648447
      MultiLine       =   0   'False
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":3025
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkLock 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   4860
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   91
      TabStop         =   0   'False
      ToolTipText     =   " Synchronise Scroll Bars "
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin VB.CheckBox chkLock 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   1
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   90
      TabStop         =   0   'False
      ToolTipText     =   " Synchronise Scroll Bars "
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin VB.CheckBox chkLock 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   4260
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   89
      TabStop         =   0   'False
      ToolTipText     =   " Synchronise Scroll Bars "
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4890
      TabIndex        =   67
      TabStop         =   0   'False
      ToolTipText     =   " Reset adjustment factor "
      Top             =   6900
      Width           =   195
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4890
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   " Reset interval size  "
      Top             =   6645
      Width           =   195
   End
   Begin VB.HScrollBar ColorGradient 
      Height          =   195
      Index           =   1
      LargeChange     =   55
      Left            =   1440
      Max             =   750
      Min             =   250
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   6900
      Value           =   493
      Width           =   3400
   End
   Begin VB.HScrollBar ColorGradient 
      Height          =   195
      Index           =   0
      LargeChange     =   3
      Left            =   1440
      Max             =   30
      Min             =   1
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   6645
      Value           =   15
      Width           =   3400
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   0
      Left            =   2340
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   " CTRL+C to copy to clipboard "
      Top             =   420
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":30A0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.VScrollBar ColorSelect 
      Height          =   4635
      Index           =   2
      LargeChange     =   10
      Left            =   4860
      Max             =   0
      Min             =   255
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1200
      Width           =   240
   End
   Begin VB.VScrollBar ColorSelect 
      Height          =   4635
      Index           =   1
      LargeChange     =   10
      Left            =   4560
      Max             =   0
      Min             =   255
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   240
   End
   Begin VB.VScrollBar ColorSelect 
      Height          =   4635
      Index           =   0
      LargeChange     =   10
      Left            =   4260
      Max             =   0
      Min             =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   240
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   2
      Left            =   2340
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1500
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   15332340
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":311B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   3
      Left            =   2340
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   15332340
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":3196
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   4
      Left            =   2340
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2580
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   15332340
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":3211
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   5
      Left            =   2340
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   15332340
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":328C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   6
      Left            =   2340
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3660
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   15332340
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":3307
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   7
      Left            =   2340
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   15332340
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":3382
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   8
      Left            =   2340
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4740
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   15332340
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":33FD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   9
      Left            =   2340
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5280
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   15332340
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":3478
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   1
      Left            =   2340
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   15332340
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":34F3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtNewDec 
      Height          =   255
      Index           =   1
      Left            =   3420
      TabIndex        =   2
      ToolTipText     =   " Enter decimal value for green "
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      MaxLength       =   3
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":356E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtNewDec 
      Height          =   255
      Index           =   2
      Left            =   3780
      TabIndex        =   3
      ToolTipText     =   " Enter decimal value for blue "
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      MaxLength       =   3
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":35E9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtColorCode 
      Height          =   255
      Index           =   10
      Left            =   2340
      TabIndex        =   243
      TabStop         =   0   'False
      Top             =   5820
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   15332340
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      MaxLength       =   7
      Appearance      =   0
      TextRTF         =   $"ColorSelect.frx":3664
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblClipBoard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   3420
      TabIndex        =   262
      Top             =   6240
      Width           =   1700
   End
   Begin VB.Label lblClipBoard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Index           =   0
      Left            =   2340
      TabIndex        =   251
      ToolTipText     =   " Highlight destination field and press CTRL+V "
      Top             =   6240
      Width           =   1100
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E9F3F4&
      Caption         =   "Hex value in Clipboard:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   250
      Top             =   6270
      Width           =   2100
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   11
      Left            =   2205
      TabIndex        =   249
      Top             =   5850
      Width           =   120
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   3060
      TabIndex        =   248
      Top             =   5820
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   3420
      TabIndex        =   247
      Top             =   5820
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   3780
      TabIndex        =   246
      Top             =   5820
      Width           =   375
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   10
      Left            =   90
      TabIndex        =   245
      ToolTipText     =   " Click to select "
      Top             =   5520
      Width           =   1845
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   1920
      TabIndex        =   244
      ToolTipText     =   " Click to add to final scheme "
      Top             =   5820
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   27
      Left            =   4890
      TabIndex        =   115
      Top             =   1050
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   26
      Left            =   4890
      TabIndex        =   114
      Top             =   900
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   25
      Left            =   4890
      TabIndex        =   113
      Top             =   780
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   24
      Left            =   4890
      TabIndex        =   112
      Top             =   660
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   23
      Left            =   4890
      TabIndex        =   111
      Top             =   540
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   22
      Left            =   4890
      TabIndex        =   110
      Top             =   420
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   21
      Left            =   4890
      TabIndex        =   109
      Top             =   300
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   20
      Left            =   4890
      TabIndex        =   108
      Top             =   180
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   19
      Left            =   4890
      TabIndex        =   107
      Top             =   60
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   4290
      TabIndex        =   106
      Top             =   300
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   18
      Left            =   4590
      TabIndex        =   105
      Top             =   1050
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   17
      Left            =   4590
      TabIndex        =   104
      Top             =   900
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   16
      Left            =   4590
      TabIndex        =   103
      Top             =   780
      Width           =   195
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   10
      Left            =   2205
      TabIndex        =   102
      Top             =   5310
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   9
      Left            =   2205
      TabIndex        =   101
      Top             =   4770
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   8
      Left            =   2205
      TabIndex        =   100
      Top             =   4230
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   7
      Left            =   2205
      TabIndex        =   99
      Top             =   3690
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   6
      Left            =   2205
      TabIndex        =   98
      Top             =   3150
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   5
      Left            =   2205
      TabIndex        =   97
      Top             =   2610
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   4
      Left            =   2205
      TabIndex        =   96
      Top             =   2070
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   3
      Left            =   2205
      TabIndex        =   95
      Top             =   1530
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   2
      Left            =   2205
      TabIndex        =   94
      Top             =   990
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   1
      Left            =   2205
      TabIndex        =   93
      Top             =   480
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E9F3F4&
      Caption         =   "#"
      Height          =   195
      Index           =   0
      Left            =   2200
      TabIndex        =   92
      Top             =   150
      Width           =   120
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   14
      Left            =   4590
      TabIndex        =   88
      Top             =   540
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   9
      Left            =   4290
      TabIndex        =   87
      Top             =   1050
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   4
      Left            =   4290
      TabIndex        =   86
      Top             =   420
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   4290
      TabIndex        =   85
      Top             =   60
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   6
      Left            =   4290
      TabIndex        =   84
      Top             =   660
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   11
      Left            =   4590
      TabIndex        =   83
      Top             =   180
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   4290
      TabIndex        =   82
      Top             =   180
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   7
      Left            =   4290
      TabIndex        =   81
      Top             =   780
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   12
      Left            =   4590
      TabIndex        =   80
      Top             =   300
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   8
      Left            =   4290
      TabIndex        =   79
      Top             =   900
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   13
      Left            =   4590
      TabIndex        =   78
      Top             =   420
      Width           =   195
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   1920
      TabIndex        =   77
      ToolTipText     =   " Click to add to final scheme "
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   76
      ToolTipText     =   " Click to add to final scheme "
      Top             =   4740
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   75
      ToolTipText     =   " Click to add to final scheme "
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   74
      ToolTipText     =   " Click to add to final scheme "
      Top             =   3660
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   73
      ToolTipText     =   " Click to add to final scheme "
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   72
      ToolTipText     =   " Click to add to final scheme "
      Top             =   2580
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   71
      ToolTipText     =   " Click to add to final scheme "
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   70
      ToolTipText     =   " Click to add to final scheme "
      Top             =   1500
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   69
      ToolTipText     =   " Click to add to final scheme "
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   68
      ToolTipText     =   " Click to add to final scheme "
      Top             =   420
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E9F3F4&
      Caption         =   "Interval ="
      Height          =   195
      Left            =   180
      TabIndex        =   65
      Top             =   6645
      Width           =   780
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E9F3F4&
      Caption         =   "Factor ="
      Height          =   195
      Left            =   180
      TabIndex        =   64
      Top             =   6900
      Width           =   780
   End
   Begin VB.Label lblInterval 
      BackColor       =   &H00E9F3F4&
      Height          =   195
      Left            =   1020
      TabIndex        =   63
      Top             =   6645
      Width           =   420
   End
   Begin VB.Label lblFactor 
      BackColor       =   &H00E9F3F4&
      Height          =   195
      Left            =   1020
      TabIndex        =   61
      Top             =   6900
      Width           =   420
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   15
      Left            =   4590
      TabIndex        =   59
      Top             =   660
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   10
      Left            =   4590
      TabIndex        =   58
      Top             =   60
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   5
      Left            =   4290
      TabIndex        =   57
      Top             =   540
      Width           =   195
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   9
      Left            =   90
      TabIndex        =   55
      ToolTipText     =   " Click to select "
      Top             =   4980
      Width           =   1845
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   8
      Left            =   90
      TabIndex        =   54
      ToolTipText     =   " Click to select "
      Top             =   4440
      Width           =   1845
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   3780
      TabIndex        =   53
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   3420
      TabIndex        =   52
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   3060
      TabIndex        =   51
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   3780
      TabIndex        =   49
      Top             =   4740
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   3420
      TabIndex        =   48
      Top             =   4740
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   3060
      TabIndex        =   47
      Top             =   4740
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3780
      TabIndex        =   45
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3420
      TabIndex        =   44
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3060
      TabIndex        =   43
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3780
      TabIndex        =   41
      Top             =   3660
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3420
      TabIndex        =   40
      Top             =   3660
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3060
      TabIndex        =   39
      Top             =   3660
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3780
      TabIndex        =   37
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3420
      TabIndex        =   36
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3060
      TabIndex        =   35
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3780
      TabIndex        =   33
      Top             =   2580
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3420
      TabIndex        =   32
      Top             =   2580
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3060
      TabIndex        =   31
      Top             =   2580
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3780
      TabIndex        =   29
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3420
      TabIndex        =   28
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3060
      TabIndex        =   27
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3780
      TabIndex        =   25
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3420
      TabIndex        =   24
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3060
      TabIndex        =   23
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3780
      TabIndex        =   21
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3420
      TabIndex        =   20
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3060
      TabIndex        =   19
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   3780
      TabIndex        =   18
      ToolTipText     =   " CTRL+C to copy to clipboard "
      Top             =   420
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   3420
      TabIndex        =   17
      ToolTipText     =   " CTRL+C to copy to clipboard "
      Top             =   420
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9F3F4&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   3060
      TabIndex        =   16
      ToolTipText     =   " CTRL+C to copy to clipboard "
      Top             =   420
      Width           =   375
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   7
      Left            =   90
      TabIndex        =   14
      ToolTipText     =   " Click to select "
      Top             =   3900
      Width           =   1845
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   6
      Left            =   90
      TabIndex        =   13
      ToolTipText     =   " Click to select "
      Top             =   3360
      Width           =   1845
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   5
      Left            =   90
      TabIndex        =   12
      ToolTipText     =   " Click to select "
      Top             =   2820
      Width           =   1845
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   4
      Left            =   90
      TabIndex        =   11
      ToolTipText     =   " Click to select "
      Top             =   2280
      Width           =   1845
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   3
      Left            =   90
      TabIndex        =   10
      ToolTipText     =   " Click to select "
      Top             =   1740
      Width           =   1845
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   2
      Left            =   90
      TabIndex        =   9
      ToolTipText     =   " Click to select "
      Top             =   1200
      Width           =   1845
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   1
      Left            =   90
      TabIndex        =   8
      ToolTipText     =   " Click to select "
      Top             =   660
      Width           =   1845
   End
   Begin VB.Label lblSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   0
      Left            =   90
      TabIndex        =   7
      ToolTipText     =   " Click to select "
      Top             =   120
      Width           =   1845
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu RecentFileItem 
         Caption         =   "Open Color File"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu RecentFileItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu RecentFileItem 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu RecentFileItem 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu RecentFileItem 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu RecentFileItem 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu RecentFileItem 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu RecentFileItem 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu RecentFileItem 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu RecentFileItem 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu RecentFileItem 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu FileItem 
         Caption         =   "Save"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu FileItem 
         Caption         =   "Save As"
         Index           =   2
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu FileItem 
         Caption         =   "Exit"
         Index           =   4
      End
   End
   Begin VB.Menu mView 
      Caption         =   "View"
      WindowList      =   -1  'True
      Begin VB.Menu ViewItem 
         Caption         =   "Show Palette"
         Index           =   0
         Shortcut        =   {F4}
      End
      Begin VB.Menu ViewItem 
         Caption         =   "Hide Palette"
         Index           =   1
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mPref 
      Caption         =   "Preferences"
   End
   Begin VB.Menu mAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IndexFore_M      As Long
Private IndexBack_M      As Long

Private Function InvertHexValue(ByVal vHex As String) As String

On Error GoTo errhandler

    vHex = Trim$(vHex)
    
    If Len(vHex) < 6 Then
        vHex = String$(Len(vHex) - 6, "0") & vHex
    End If
    
    InvertHexValue = Right$(vHex, 2) & Mid$(vHex, 3, 2) & Left$(vHex, 2)
        
    Exit Function
        
errhandler:
    InvertHexValue = vbNullString
    Exit Function
End Function


Public Function GetOpenFilePath() As String

On Error GoTo errhandler

    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Open Color Scheme"
    CommonDialog1.InitDir = dir_OpenFolder
    CommonDialog1.FileName = dir_OpenFileTitle
    CommonDialog1.Filter = "All files (*.*)|*.*|Color scheme (*.csc)|*.csc"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    
    GetOpenFilePath = CommonDialog1.FileName
    
    Exit Function
    
errhandler:
    GetOpenFilePath = vbNullString
End Function
Public Sub FileSave(ByVal FilePath As String)

On Error GoTo errhandler

Dim N                   As Long
Dim fil                 As Long

    fil = FreeFile
    Open LCase$(FilePath) For Output As fil
    
        Print #fil, "[GENERAL SETTINGS]"                            ' [GENERAL SETTINGS]
        Print #fil, ColorSelect(0).Value
        Print #fil, ColorSelect(1).Value
        Print #fil, ColorSelect(2).Value
        
        Print #fil, ColorGradient(0).Value
        Print #fil, ColorGradient(1).Value
        
        Print #fil, lblUsed(0).Visible
        Print #fil, lblUsed(1).Visible
        Print #fil, lblUsed(2).Visible
        Print #fil, lblUsed(3).Visible
        Print #fil, lblUsed(4).Visible
        Print #fil, lblUsed(5).Visible
        Print #fil, lblUsed(6).Visible
        Print #fil, lblUsed(7).Visible
        Print #fil, lblUsed(8).Visible
        Print #fil, lblUsed(9).Visible
        
        Print #fil, txtNewDec(0).Text
        Print #fil, txtNewDec(1).Text
        Print #fil, txtNewDec(2).Text
        
        Print #fil, txtNewHexValue.Text
        
        Print #fil, " "                                             ' [FINAL COLORS 1 - 10]
        Print #fil, "[FINAL COLORS 1 - 10]"
        For N = 1 To 10
            Print #fil, lblFinal(N).Visible
            Print #fil, lblFinalNo(N).Visible
            Print #fil, lblFinal(N).BackColor
            Print #fil, lblFinal(N).ToolTipText
            Print #fil, txtFinal(N).Visible
            Print #fil, txtFinal(N + 10).Visible
            Print #fil, txtFinal(N).Text
            Print #fil, txtFinal(N + 10).Text
        Next N
                
        Print #fil, " "
        Print #fil, "[PREFERENCES]"
        Print #fil, pref_UserName
        pref_Description = Replace(pref_Description, vbCrLf, "@@")
        Print #fil, pref_Description
        Print #fil, pref_Initials
        
        Print #fil, " "
        Print #fil, "[LINKS]"
        For N = 1 To 10
            Print #fil, txtLinkForeColorHex(N).Text
            Print #fil, txtLinkForeColor(N).Text
            Print #fil, txtLinkBackColorHex(N).Text
            Print #fil, txtLinkBackColor(N).Text
            Print #fil, lblLinkColor(N).ForeColor
            Print #fil, lblLinkColor(N).BackColor
        Next N
        
        Print #fil, " "
        Print #fil, "[TEXT]"
        For N = 1 To 5
            Print #fil, txtForeColorHex(N).Text
            Print #fil, txtForeColor(N).Text
            Print #fil, txtBackColorHex(N).Text
            Print #fil, txtBackColor(N).Text
            Print #fil, lblTextColor(N).ForeColor
            Print #fil, lblTextColor(N).BackColor
        Next N
        
    Close fil
    
    ' clear file_dirty flag
    File_Dirty_G = False
    
    dir_SaveFolder = LCase$(GetDirectoryPath(FilePath))
    dir_SaveFileTitle = LCase$(GetFileTitleFromPath(FilePath))
    
    Me.Caption = "ColorMate ver." & App.Major & "." & App.Minor & "." & App.Revision & "    [ " & UCase$(dir_SaveFileTitle) & " ]"
    AddToLastOpened LCase$(FilePath)
    
    Exit Sub
    
errhandler:
    ' set file_dirty flag
    File_Dirty_G = True
    Exit Sub
End Sub

Public Sub AddToLastOpened(ByVal FilePath As String)
    
On Error GoTo errhandler

Dim N                   As Long
Dim P                   As Long
Dim count               As String

ReDim x(1 To 10) As String
ReDim y(1 To 9) As String
    
    FilePath = LCase$(FilePath)
    
    ' store new path at position 1 in X()
    If swFileExists(FilePath) Then
        x(1) = FilePath
        For N = 2 To 10
            If swFileExists(RecentFileItem(N).Caption) Then
                x(N) = RecentFileItem(N).Caption
            End If
        Next N
        ReDim Preserve x(1 To 9)
    Else
        For N = 2 To 10
            If swFileExists(RecentFileItem(N).Caption) Then
                x(N - 1) = RecentFileItem(N).Caption
            End If
        Next N
    End If
        
    ' add paths from X() to Y()
    count = 0
    For N = 1 To 9
        If Len(x(N)) = 0 Then GoTo NextN
        For P = 1 To 9
            If y(P) = x(N) And Len(x(N)) > 0 Then
                GoTo NextN
            End If
        Next P
        If P = 10 And swFileExists(x(N)) And InStr(x(N), ":") > 0 Then
            count = count + 1
            y(count) = x(N)
        End If
NextN:
    Next N
    
    If count > 0 Then
        ReDim Preserve y(1 To count)
    Else
        ' hide all submenu items (2 - 10)
        For N = 2 To 10
            RecentFileItem(N).Caption = vbNullString
            RecentFileItem(N).Visible = False
        Next N
        GoTo errhandler
    End If
    
    ' hide all submenu items (2 - 10)
    For N = 2 To 10
            RecentFileItem(N).Caption = vbNullString
            RecentFileItem(N).Visible = False
        Next N
        
    ' load RecentFileItem (2 - 10) array from Y()
    For N = LBound(y) To UBound(y)
        If Len(y(N)) > 0 Then
            RecentFileItem(N + 1).Visible = True
            RecentFileItem(N + 1).Caption = y(N)
        End If
    Next N
        
errhandler:
    Exit Sub
End Sub

Public Sub FileOpen(ByVal FilePath As String)

On Error Resume Next

Dim N                   As Long
Dim fil                 As Long
Dim tmp                 As String
    
    pref_UserName = vbNullString
    pref_Description = vbNullString
    pref_Initials = vbNullString
    
    For N = 1 To 10
        lblTarget(N).Visible = False
    Next N
    
    fil = FreeFile
    Open FilePath For Input As fil
        
        Line Input #fil, tmp                                        ' [GENERAL SETTINGS]
        Line Input #fil, tmp: ColorSelect(0).Value = CInt(tmp)
        Line Input #fil, tmp: ColorSelect(1).Value = CInt(tmp)
        Line Input #fil, tmp: ColorSelect(2).Value = CInt(tmp)
        
        Line Input #fil, tmp: ColorGradient(0).Value = CInt(tmp)
        Line Input #fil, tmp: ColorGradient(1).Value = CInt(tmp)
        
        Line Input #fil, tmp: lblUsed(0).Visible = CBool(tmp)
        Line Input #fil, tmp: lblUsed(1).Visible = CBool(tmp)
        Line Input #fil, tmp: lblUsed(2).Visible = CBool(tmp)
        Line Input #fil, tmp: lblUsed(3).Visible = CBool(tmp)
        Line Input #fil, tmp: lblUsed(4).Visible = CBool(tmp)
        Line Input #fil, tmp: lblUsed(5).Visible = CBool(tmp)
        Line Input #fil, tmp: lblUsed(6).Visible = CBool(tmp)
        Line Input #fil, tmp: lblUsed(7).Visible = CBool(tmp)
        Line Input #fil, tmp: lblUsed(8).Visible = CBool(tmp)
        Line Input #fil, tmp: lblUsed(9).Visible = CBool(tmp)
        
        Line Input #fil, tmp: txtNewDec(0).Text = tmp
        Line Input #fil, tmp: txtNewDec(1).Text = tmp
        Line Input #fil, tmp: txtNewDec(2).Text = tmp
        
        Line Input #fil, tmp: txtNewHexValue.Text = tmp
        
        Line Input #fil, tmp                                        ' [FINAL COLORS 1 - 10]
        Line Input #fil, tmp
        For N = 1 To 10
            Line Input #fil, tmp: lblFinal(N).Visible = CBool(tmp)
            Line Input #fil, tmp: lblFinalNo(N).Visible = CBool(tmp)
            Line Input #fil, tmp: lblFinal(N).BackColor = CLng(tmp)
            Line Input #fil, tmp: lblFinal(N).ToolTipText = tmp
            Line Input #fil, tmp: txtFinal(N).Visible = CBool(tmp)
            Line Input #fil, tmp: txtFinal(N + 10).Visible = CBool(tmp)
            Line Input #fil, tmp: txtFinal(N).Text = tmp
            Line Input #fil, tmp: txtFinal(N + 10).Text = tmp
        Next N
                                                                    ' [PREFERENCES]
        Line Input #fil, tmp
        Line Input #fil, tmp
        Line Input #fil, pref_UserName
        Line Input #fil, pref_Description
        pref_Description = Replace(pref_Description, "@@", vbCrLf)
        Line Input #fil, pref_Initials
        
        Line Input #fil, tmp: tmp = vbNullString                    ' [LINKS]
        Line Input #fil, tmp: tmp = vbNullString
        For N = 1 To 10
            Line Input #fil, tmp: txtLinkForeColorHex(N).Text = tmp: tmp = vbNullString
            Line Input #fil, tmp: txtLinkForeColor(N).Text = tmp: tmp = vbNullString
            Line Input #fil, tmp: txtLinkBackColorHex(N).Text = tmp: tmp = vbNullString
            Line Input #fil, tmp: txtLinkBackColor(N).Text = tmp: tmp = vbNullString
            Line Input #fil, tmp: lblLinkColor(N).ForeColor = CLng(tmp): tmp = vbNullString
            Line Input #fil, tmp: lblLinkColor(N).BackColor = CLng(tmp): tmp = vbNullString
        Next N
        
        Line Input #fil, tmp                                        ' [TEXT]
        Line Input #fil, tmp
        For N = 1 To 5
            Line Input #fil, tmp:  txtForeColorHex(N).Text = tmp: tmp = vbNullString
            Line Input #fil, tmp:  txtForeColor(N).Text = tmp: tmp = vbNullString
            Line Input #fil, tmp:  txtBackColorHex(N).Text = tmp: tmp = vbNullString
            Line Input #fil, tmp:  txtBackColor(N).Text = tmp: tmp = vbNullString
            Line Input #fil, tmp:  lblTextColor(N).ForeColor = CLng(tmp): tmp = vbNullString
            Line Input #fil, tmp:  lblTextColor(N).BackColor = CLng(tmp): tmp = vbNullString
        Next N
        
    Close fil
    
    ' open-path and filename is stored in registry and
    ' used as default next time the program is launched
    dir_OpenFolder = LCase$(GetDirectoryPath(FilePath))
    dir_OpenFileTitle = LCase$(GetFileTitleFromPath(FilePath))
    
    ' set save folder and file title to that of the  file just opened
    dir_SaveFolder = dir_OpenFolder
    dir_SaveFileTitle = dir_OpenFileTitle
    
    ' update form caption with new filename
    Me.Caption = "ColorMate ver." & App.Major & "." & App.Minor & "." & App.Revision & "    [ " & UCase$(dir_OpenFileTitle) & " ]"
    ColorSelect_Scroll 0
    
    AddToLastOpened LCase$(FilePath)
    
    ' clear file_dirty flag
    File_Dirty_G = False
    
errhandler:
    Close fil
    Exit Sub
End Sub

















Public Function GetSaveFilePath() As String
    
On Error GoTo errhandler
        
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Save Color Scheme"
    CommonDialog1.InitDir = dir_SaveFolder
    CommonDialog1.FileName = dir_SaveFileTitle
    CommonDialog1.Filter = "All files (*.*)|*.*|Color scheme (*.csc)|*.csc"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.DefaultExt = "csc"
    CommonDialog1.Flags = &H2
    CommonDialog1.ShowSave
    
    GetSaveFilePath = CommonDialog1.FileName
    
    Exit Function
    
errhandler:
    GetSaveFilePath = vbNullString
End Function

'------------------------------------------------------------------------------
' red =
' green =
' blue =
' grey =
'------------------------------------------------------------------------------
Private Sub OverFlowWarning()
    
On Error GoTo errhandler

Dim N                       As Long
    
Const red = &HC0C0FF
Const green = &HC0FFC0
Const blue = &HFFC0C0
Const grey = &H8000000F
Const white = &HFFFFFF

    For N = 0 To 10
        lblRed(N).BackColor = &HE9F3F4
        lblGreen(N).BackColor = &HE9F3F4
        lblBlue(N).BackColor = &HE9F3F4
        txtColorCode(N).BackColor = &HFFFFFF '&HE9F3F4
    Next N
    
    For N = 0 To 10
        Select Case Val(lblRed(N).Caption)
            Case 0, 255
                lblRed(N).BackColor = &HDEE1FF
                txtColorCode(N).BackColor = &HE9F3F4
        End Select
        
        Select Case Val(lblGreen(N).Caption)
            Case 0, 255
                lblGreen(N).BackColor = &HD3FFE4
                txtColorCode(N).BackColor = &HE9F3F4
        End Select
        
        Select Case Val(lblBlue(N).Caption)
            Case 0, 255
                lblBlue(N).BackColor = &HEFE3DC
                txtColorCode(N).BackColor = &HE9F3F4
        End Select
    Next N
    
errhandler:
    Exit Sub
End Sub

Private Sub Select_General_Color(ByVal Index As Long)

On Error GoTo errhandler

Dim N                   As Long
Dim P                   As Long
Dim FirstFree           As Long
    
    ' set file_dirty flag
    File_Dirty_G = True

    ' find first free position for Final color
    For N = 1 To 10
        If lblFinal(N).Visible = False Then
            FirstFree = N
            Exit For
        End If
    Next N
    
    ' if all 10 positions are occupied then tell user
    If FirstFree = 0 Then
        StatusBar1.Panels.Item(1).Text = "Maximum colors is 10!  Shift-Click a color field to clear it."
        GoTo errhandler
    End If
    
    ' check if the color is already selected
    For N = 1 To 10
        If lblFinal(N).BackColor = lblUsed(Index).BackColor Then
            GoTo errhandler
        End If
    Next N
    
    ' display free field
    lblFinal(FirstFree).Visible = True
    lblFinalNo(FirstFree).Visible = True
    txtFinal(FirstFree).Visible = True
    txtFinal(FirstFree + 10).Visible = True
    
    ' enter color and name label for selected Final color
    lblFinal(FirstFree).BackColor = lblUsed(Index).BackColor
    txtFinal(FirstFree).Text = " " & txtColorCode(Index).Text
    txtFinal(FirstFree + 10) = " " & Format$(lblRed(Index).Caption, "000") & ", " & _
                                     Format$(lblGreen(Index).Caption, "000") & ", " & _
                                     Format$(lblBlue(Index).Caption, "000")
    
errhandler:
    Exit Sub
End Sub

Private Sub Select_Link_Color(ByVal Index As Long)

On Error GoTo errhandler

Dim N                   As Long
Dim SelItem             As Long
Dim IsForeground        As Boolean
    
    If IndexFore_M > 0 Then
        SelItem = IndexFore_M
        IsForeground = True
        
    ElseIf IndexBack_M > 0 Then
        SelItem = IndexBack_M
        IsForeground = False
        
    Else
        Exit Sub
    End If
        
    ' set file_dirty flag
    File_Dirty_G = True

    If IsForeground Then
        lblLinkColor(SelItem).ForeColor = lblUsed(Index).BackColor
        txtLinkForeColorHex(SelItem).Text = txtColorCode(Index).Text
        txtLinkForeColor(SelItem).Text = " " & Format$(lblRed(Index).Caption, "000") & ", " & _
                                               Format$(lblGreen(Index).Caption, "000") & ", " & _
                                               Format$(lblBlue(Index).Caption, "000")
    Else
        lblLinkColor(SelItem).BackColor = lblUsed(Index).BackColor
        txtLinkBackColorHex(SelItem).Text = txtColorCode(Index).Text
        txtLinkBackColor(SelItem).Text = " " & Format$(lblRed(Index).Caption, "000") & ", " & _
                                               Format$(lblGreen(Index).Caption, "000") & ", " & _
                                               Format$(lblBlue(Index).Caption, "000")
    End If
       
    
errhandler:
    Exit Sub
End Sub

Private Sub Select_Text_Color(ByVal Index As Long)

On Error GoTo errhandler

Dim N                   As Long
Dim SelItem             As Long
Dim IsForeground        As Boolean
    
    If IndexFore_M > 0 Then
        SelItem = IndexFore_M
        IsForeground = True
        
    ElseIf IndexBack_M > 0 Then
        SelItem = IndexBack_M
        IsForeground = False
        
    Else
        Exit Sub
    End If
        
    ' set file_dirty flag
    File_Dirty_G = True

    If IsForeground Then
        lblTextColor(SelItem).ForeColor = lblUsed(Index).BackColor
        txtForeColorHex(SelItem).Text = txtColorCode(Index).Text
        txtForeColor(SelItem).Text = " " & Format$(lblRed(Index).Caption, "000") & ", " & _
                                           Format$(lblGreen(Index).Caption, "000") & ", " & _
                                           Format$(lblBlue(Index).Caption, "000")
    Else
        lblTextColor(SelItem).BackColor = lblUsed(Index).BackColor
        txtBackColorHex(SelItem).Text = txtColorCode(Index).Text
        txtBackColor(SelItem).Text = " " & Format$(lblRed(Index).Caption, "000") & ", " & _
                                           Format$(lblGreen(Index).Caption, "000") & ", " & _
                                           Format$(lblBlue(Index).Caption, "000")
    End If
       
    
errhandler:
    Exit Sub
End Sub


Private Sub chkLock_Click(Index As Integer)

On Error GoTo errhandler

    ' code to synchronise scroll bars
    ColorSelect(0).Enabled = True
    ColorSelect(1).Enabled = True
    ColorSelect(2).Enabled = True
    
    If chkLock(0).Value = 1 And _
       chkLock(1).Value = 1 And _
       chkLock(2).Value = 1 Then            ' R G B
       
        ColorSelect(0).Enabled = False
        ColorSelect(1).Enabled = False
        ColorSelect(2).Enabled = True
        
    ElseIf chkLock(0).Value = 1 And _
           chkLock(1).Value = 1 And _
           chkLock(2).Value = 0 Then        ' R G b
           
        ColorSelect(0).Enabled = False
        ColorSelect(1).Enabled = True
        ColorSelect(2).Enabled = True
        
    ElseIf chkLock(0).Value = 0 And _
           chkLock(1).Value = 1 And _
           chkLock(2).Value = 1 Then        ' r G B
           
        ColorSelect(0).Enabled = True
        ColorSelect(1).Enabled = False
        ColorSelect(2).Enabled = True
        
    ElseIf chkLock(0).Value = 1 And _
           chkLock(1).Value = 0 And _
           chkLock(2).Value = 1 Then        ' R g B
           
        ColorSelect(0).Enabled = False
        ColorSelect(1).Enabled = True
        ColorSelect(2).Enabled = True
        
    Else
        ColorSelect(0).Enabled = True       ' r g b
        ColorSelect(1).Enabled = True
        ColorSelect(2).Enabled = True
    End If
        
errhandler:
    Exit Sub
End Sub




Public Sub cmdReset_Click(Index As Integer)
    
On Error GoTo errhandler
    
    ' reset interval and adjustment factor settings
    Select Case Index
        Case 0
            ColorGradient(0).Value = 15
            int_Gradient_Interval = ColorGradient(0).Value
            lblInterval.Caption = int_Gradient_Interval
            
        Case 1
            ColorGradient(1).Value = 494
            int_Gradient_Factor = ColorGradient(1).Value
            sng_Gradient_Factor = 494 / ColorGradient(1).Value
            lblFactor.Caption = Format$(494 / ColorGradient(1).Value, "0.00")
            
    End Select
    
    ColorSelect_Scroll 0
    
errhandler:
    Exit Sub
End Sub

Private Sub ColorGradient_Change(Index As Integer)
    
On Error GoTo errhandler

Dim N                   As Long
    
    ' update scroll bars for interval and factor adjustment
    int_Gradient_Interval = ColorGradient(0).Value
    lblInterval.Caption = int_Gradient_Interval
    
    int_Gradient_Factor = ColorGradient(1).Value
    sng_Gradient_Factor = 494 / ColorGradient(1).Value
    lblFactor.Caption = Format$(494 / ColorGradient(1).Value, "0.00")
        
    ColorSelect_Scroll 0
        
errhandler:
    Exit Sub
End Sub

Private Sub ColorSelect_Change(Index As Integer)

On Error GoTo errhandler

Dim N                   As Long
Dim vRed                As Integer
Dim vGreen              As Integer
Dim vBlue               As Integer
Dim hexRed              As String
Dim hexGreen            As String
Dim hexBlue             As String
            
    ' synchronise scroll bars depending on lock check box settings
    If chkLock(0).Value = 1 And chkLock(1).Value = 1 And chkLock(2).Value = 1 Then
        ColorSelect(0).Value = ColorSelect(2).Value
        ColorSelect(1).Value = ColorSelect(2).Value
    ElseIf chkLock(0).Value = 1 And chkLock(1).Value = 1 Then
        ColorSelect(0).Value = ColorSelect(1).Value
    ElseIf chkLock(0).Value = 1 And chkLock(2).Value = 1 Then
        ColorSelect(0).Value = ColorSelect(2).Value
    ElseIf chkLock(1).Value = 1 And chkLock(2).Value = 1 Then
        ColorSelect(1).Value = ColorSelect(2).Value
    End If
        
    ' process R,G,B values for master color
    vRed = ColorSelect(0).Value
    If vRed > 255 Then vRed = 255
    If vRed < 0 Then vRed = 0
    hexRed = Hex(vRed)
    If Len(hexRed) = 1 Then hexRed = "0" & hexRed
    lblRed(0).Caption = Format$(vRed, "000")
    
    vGreen = ColorSelect(1).Value
    If vGreen > 255 Then vGreen = 255
    If vGreen < 0 Then vGreen = 0
    hexGreen = Hex(vGreen)
    If Len(hexGreen) = 1 Then hexGreen = "0" & hexGreen
    lblGreen(0).Caption = Format$(vGreen, "000")
    
    vBlue = ColorSelect(2).Value
    If vBlue > 255 Then vBlue = 255
    If vBlue < 0 Then vBlue = 0
    hexBlue = Hex(vBlue)
    If Len(hexBlue) = 1 Then hexBlue = "0" & hexBlue
    lblBlue(0).Caption = Format$(vBlue, "000")
        
    ' process R,G,B values for derived colors
    For N = 0 To 10
    
        vRed = CInt((ColorSelect(0).Value - N * int_Gradient_Interval) * sng_Gradient_Factor)
        If vRed > 255 Then vRed = 255
        If vRed < 0 Then vRed = 0
        hexRed = Hex(vRed)
        If Len(hexRed) = 1 Then hexRed = "0" & hexRed
        lblRed(N).Caption = Format$(vRed, "000")
        
        vGreen = CInt((ColorSelect(1).Value - N * int_Gradient_Interval) * sng_Gradient_Factor)
        If vGreen > 255 Then vGreen = 255
        If vGreen < 0 Then vGreen = 0
        hexGreen = Hex(vGreen)
        If Len(hexGreen) = 1 Then hexGreen = "0" & hexGreen
        lblGreen(N).Caption = Format$(vGreen, "000")
        
        vBlue = CInt((ColorSelect(2).Value - N * int_Gradient_Interval) * sng_Gradient_Factor)
        If vBlue > 255 Then vBlue = 255
        If vBlue < 0 Then vBlue = 0
        hexBlue = Hex(vBlue)
        If Len(hexBlue) = 1 Then hexBlue = "0" & hexBlue
        lblBlue(N).Caption = Format$(vBlue, "000")
        
        ' display results
        lblSample(N).BackColor = RGB(vRed, vGreen, vBlue)
        lblUsed(N).BackColor = lblSample(N).BackColor
        txtColorCode(N).Text = Space$(1) & hexRed & hexGreen & hexBlue
                
    Next N
        
errhandler:
    Exit Sub
    
End Sub

Private Sub ColorSelect_GotFocus(Index As Integer)
    
On Error GoTo errhandler

Dim N As Long
Dim count As Long

    For N = 0 To 2
        If chkLock(N).Value = 1 Then
            count = count + 1
        End If
    Next N
    If count = 1 Then
        chkLock(0).Value = 0
        chkLock(1).Value = 0
        chkLock(2).Value = 0
    End If
    
errhandler:
    Exit Sub
End Sub

Private Sub ColorSelect_Scroll(Index As Integer)

On Error GoTo errhandler

Dim N                       As Long
Dim vRed                    As Integer
Dim vGreen                  As Integer
Dim vBlue                   As Integer
Dim hexRed                  As String
Dim hexGreen                As String
Dim hexBlue                 As String
        
    If chkLock(0).Value = 1 And chkLock(1).Value = 1 And chkLock(2).Value = 1 Then
        ColorSelect(0).Value = ColorSelect(2).Value
        ColorSelect(1).Value = ColorSelect(2).Value
    ElseIf chkLock(0).Value = 1 And chkLock(1).Value = 1 Then
        ColorSelect(0).Value = ColorSelect(1).Value
    ElseIf chkLock(0).Value = 1 And chkLock(2).Value = 1 Then
        ColorSelect(0).Value = ColorSelect(2).Value
    ElseIf chkLock(1).Value = 1 And chkLock(2).Value = 1 Then
        ColorSelect(1).Value = ColorSelect(2).Value
    End If
        
    vRed = ColorSelect(0).Value
    If vRed > 255 Then vRed = 255
    If vRed < 0 Then vRed = 0
    hexRed = Hex(vRed)
    If Len(hexRed) = 1 Then hexRed = "0" & hexRed
    lblRed(0).Caption = Format$(vRed, "000")
    
    vGreen = ColorSelect(1).Value
    If vGreen > 255 Then vGreen = 255
    If vGreen < 0 Then vGreen = 0
    hexGreen = Hex(vGreen)
    If Len(hexGreen) = 1 Then hexGreen = "0" & hexGreen
    lblGreen(0).Caption = Format$(vGreen, "000")
    
    vBlue = ColorSelect(2).Value
    If vBlue > 255 Then vBlue = 255
    If vBlue < 0 Then vBlue = 0
    hexBlue = Hex(vBlue)
    If Len(hexBlue) = 1 Then hexBlue = "0" & hexBlue
    lblBlue(0).Caption = Format$(vBlue, "000")
        
    For N = 0 To 10
        vRed = CInt((ColorSelect(0).Value - N * int_Gradient_Interval) * sng_Gradient_Factor)
        If vRed > 255 Then vRed = 255
        If vRed < 0 Then vRed = 0
        hexRed = Hex(vRed)
        If Len(hexRed) = 1 Then hexRed = "0" & hexRed
        lblRed(N).Caption = Format$(vRed, "000")
        
        vGreen = CInt((ColorSelect(1).Value - N * int_Gradient_Interval) * sng_Gradient_Factor)
        If vGreen > 255 Then vGreen = 255
        If vGreen < 0 Then vGreen = 0
        hexGreen = Hex(vGreen)
        If Len(hexGreen) = 1 Then hexGreen = "0" & hexGreen
        lblGreen(N).Caption = Format$(vGreen, "000")
        
        vBlue = CInt((ColorSelect(2).Value - N * int_Gradient_Interval) * sng_Gradient_Factor)
        If vBlue > 255 Then vBlue = 255
        If vBlue < 0 Then vBlue = 0
        hexBlue = Hex(vBlue)
        If Len(hexBlue) = 1 Then hexBlue = "0" & hexBlue
        lblBlue(N).Caption = Format$(vBlue, "000")
        
        lblSample(N).BackColor = RGB(vRed, vGreen, vBlue)
        lblUsed(N).BackColor = lblSample(N).BackColor
        txtColorCode(N).Text = Space$(1) & hexRed & hexGreen & hexBlue
    Next N
        
    OverFlowWarning
    
errhandler:
    Exit Sub
    
End Sub







Private Sub ColorGradient_Scroll(Index As Integer)
    
On Error GoTo errhandler

    int_Gradient_Interval = ColorGradient(0).Value
    lblInterval.Caption = int_Gradient_Interval
    
    int_Gradient_Factor = ColorGradient(1).Value
    sng_Gradient_Factor = 500 / ColorGradient(1).Value
    lblFactor.Caption = Format$(500 / ColorGradient(1).Value, "0.00")
    
    ColorSelect_Scroll 0
        
    OverFlowWarning
    
errhandler:
    Exit Sub
End Sub


Private Sub FileItem_Click(Index As Integer)

On Error GoTo errhandler

Dim fil                 As Long
Dim tmp                 As String
Dim N                   As Long
    
    Select Case Index
        Case 1
            FileSave (dir_SaveFolder & dir_SaveFileTitle)       ' Save, no dialog box
                        
        Case 2                                                  ' Save color file As
            FileSave GetSaveFilePath
                        
        Case 4
            Unload Me
            
    End Select
    
errhandler:
    Close
    Exit Sub
End Sub

Private Sub Form_Load()

On Error GoTo errhandler

Dim N                       As Integer
Dim msg1                    As String
Dim msg2                    As String
Dim msg3                    As String
    
    Me.Caption = "ColorMate ver." & App.Major & "." & App.Minor & "." & App.Revision
    
    ' center form  on screen
    Me.Top = 0 'Screen.Height / 2 - Me.Height / 2
    Me.Left = 0 'Screen.Width / 2 - Me.Width / 2
    Me.Height = 8340
        
    msg1 = "Right-click the color field in the Stored Colors tab to enter a description of the web element(s) that uses the color." & vbCrLf & vbCrLf
    msg2 = "Click the description line of the web element in the lists above to retrieve the color information for the web element(s)." & vbCrLf & vbCrLf
    msg3 = "The HEX value of the selected color is automatically copied to the clipboard."
    Label.Caption = msg1 & msg2 & msg3
    
    Me.Width = 9825
    ViewItem(0).Checked = True
    ViewItem(1).Checked = False
    
    ' create default user folder for color files
    If Dir(App.Path & "\user", vbDirectory) = vbNullString Then
        MkDir App.Path & "\user"
    End If
    dir_Default = App.Path & "\user"
    
    ' get values from registry
    dir_OpenFileTitle = GetSetting("CSC", "User", "LastOpenFileTitle", "default.csc")
    dir_SaveFileTitle = GetSetting("CSC", "User", "LastSaveFileTitle", "default.csc")
    
    dir_OpenFolder = GetSetting("CSC", "User", "OpenFolder", App.Path & "\user\")
    dir_SaveFolder = GetSetting("CSC", "User", "SaveFolder", App.Path & "\user\")
    
    RecentFileItem(2).Caption = GetSetting("CSC", "User", "LastOpen02", "")
    RecentFileItem(3).Caption = GetSetting("CSC", "User", "LastOpen03", "")
    RecentFileItem(4).Caption = GetSetting("CSC", "User", "LastOpen04", "")
    RecentFileItem(5).Caption = GetSetting("CSC", "User", "LastOpen05", "")
    RecentFileItem(6).Caption = GetSetting("CSC", "User", "LastOpen06", "")
    RecentFileItem(7).Caption = GetSetting("CSC", "User", "LastOpen07", "")
    RecentFileItem(8).Caption = GetSetting("CSC", "User", "LastOpen08", "")
    RecentFileItem(9).Caption = GetSetting("CSC", "User", "LastOpen09", "")
    RecentFileItem(10).Caption = GetSetting("CSC", "User", "LastOpen10", "")
        
    For N = 2 To 10
        If Len(RecentFileItem(N).Caption) > 0 Then
            RecentFileItem(N).Visible = True
        End If
    Next N
    AddToLastOpened ("")
        
    ' reset interval and adjustment factor settings
    ColorGradient_Scroll 0
    ColorGradient_Scroll 1
    
    ' implement preset color schemes
    For N = 1 To 27
        lblColor_Click N
    Next N
        
    ' set default values for scroll bars
    ColorSelect(0).Value = 228
    ColorSelect(1).Value = 228
    ColorSelect(2).Value = 228
    ColorGradient_Change 0
    
    ' display common dialog form for open file
    FileOpen (dir_OpenFolder & dir_OpenFileTitle)
    Me.Caption = "ColorMate ver." & App.Major & "." & App.Minor & "." & App.Revision & "    [ " & UCase$(dir_OpenFileTitle) & " ]"
    
    ColorGradient(1).Value = 494
    
    SSTab1.Tab = 0
    
    StatusBar1.style = 0
    Me.Width = 9825
    ViewItem(0).Checked = True
    ViewItem(1).Checked = False
            
errhandler:
    Exit Sub
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
On Error GoTo errhandler
    
    ' unload all forms but frmMain
    Unload frmAbout
    Unload frmPreferences
    Unload frmColor
        
    ' display save dialog box if file is dirty
    If File_Dirty_G Then
        FileItem_Click 2
    End If

errhandler:
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
On Error GoTo errhandler

Dim N                   As Long
           
    SaveSetting "CSC", "User", "LastOpenFileTitle", dir_OpenFileTitle
    SaveSetting "CSC", "User", "LastSaveFileTitle", dir_SaveFileTitle
    
    SaveSetting "CSC", "User", "OpenFolder", dir_OpenFolder
    SaveSetting "CSC", "User", "SaveFolder", dir_SaveFolder
    
    SaveSetting "CSC", "User", "LastOpen01", RecentFileItem(1).Caption
    SaveSetting "CSC", "User", "LastOpen02", RecentFileItem(2).Caption
    SaveSetting "CSC", "User", "LastOpen03", RecentFileItem(3).Caption
    SaveSetting "CSC", "User", "LastOpen04", RecentFileItem(4).Caption
    SaveSetting "CSC", "User", "LastOpen05", RecentFileItem(5).Caption
    SaveSetting "CSC", "User", "LastOpen06", RecentFileItem(6).Caption
    SaveSetting "CSC", "User", "LastOpen07", RecentFileItem(7).Caption
    SaveSetting "CSC", "User", "LastOpen08", RecentFileItem(8).Caption
    SaveSetting "CSC", "User", "LastOpen09", RecentFileItem(9).Caption
    SaveSetting "CSC", "User", "LastOpen10", RecentFileItem(10).Caption
           
    ' end seqtools by closing all forms
    For N = Forms.count - 1 To 0 Step -1
        Unload Forms(N)
        DoEvents
    Next N
        
    ' prevents hanging, loaded forms
    If Forms.count > 0 Then
        End
    End If
    
errhandler:
    Exit Sub
End Sub

Private Sub lblColor_Click(Index As Integer)
    
On Error GoTo errhandler

    ' reset check boxes for locking scroll bars
    chkLock(0).Value = 0
    chkLock(1).Value = 0
    chkLock(2).Value = 0
        
    ' 15 preset color schemes
    Select Case Index
        Case 1
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 0
            lblColor(1).BackColor = RGB(255, 255, 0)
        Case 2
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 156
            lblColor(2).BackColor = RGB(255, 255, 156)
        Case 3
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 210
            lblColor(3).BackColor = RGB(255, 255, 210)
        Case 4
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 235
            lblColor(4).BackColor = RGB(255, 255, 235)          ' Yellow
        '-------------------------------------------------------
        Case 5
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 0
            ColorSelect(2).Value = 255
            lblColor(5).BackColor = RGB(255, 0, 255)
        Case 6
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 156
            ColorSelect(2).Value = 255
            lblColor(6).BackColor = RGB(255, 156, 255)
        Case 7
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 210
            ColorSelect(2).Value = 255
            lblColor(7).BackColor = RGB(255, 210, 255)
        Case 8
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 235
            ColorSelect(2).Value = 255
            lblColor(8).BackColor = RGB(255, 235, 255)          ' Magenta
        '-------------------------------------------------------
        Case 9
            ColorSelect(0).Value = 210
            ColorSelect(1).Value = 210
            ColorSelect(2).Value = 210
            lblColor(9).BackColor = RGB(210, 210, 210)          ' Grey
        '=======================================================
        Case 10
            ColorSelect(0).Value = 0
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 255
            lblColor(10).BackColor = RGB(0, 255, 255)
        Case 11
            ColorSelect(0).Value = 156
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 255
            lblColor(11).BackColor = RGB(156, 255, 255)
        Case 12
            ColorSelect(0).Value = 210
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 255
            lblColor(12).BackColor = RGB(210, 255, 255)
        Case 13
            ColorSelect(0).Value = 235
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 255
            lblColor(13).BackColor = RGB(235, 255, 255)         ' Cyan
        '-------------------------------------------------------
        Case 14
            ColorSelect(0).Value = 0
            ColorSelect(1).Value = 0
            ColorSelect(2).Value = 255
            lblColor(14).BackColor = RGB(0, 0, 255)
        Case 15
            ColorSelect(0).Value = 156
            ColorSelect(1).Value = 156
            ColorSelect(2).Value = 255
            lblColor(15).BackColor = RGB(156, 156, 255)
        Case 16
            ColorSelect(0).Value = 210
            ColorSelect(1).Value = 210
            ColorSelect(2).Value = 255
            lblColor(16).BackColor = RGB(210, 210, 255)
        Case 17
            ColorSelect(0).Value = 235
            ColorSelect(1).Value = 235
            ColorSelect(2).Value = 255
            lblColor(17).BackColor = RGB(235, 235, 255)         ' Blue
        '-------------------------------------------------------
        Case 18
            ColorSelect(0).Value = 235
            ColorSelect(1).Value = 235
            ColorSelect(2).Value = 235
            lblColor(18).BackColor = RGB(235, 235, 235)         ' Grey
        '=======================================================
        Case 19
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 0
            ColorSelect(2).Value = 0
            lblColor(19).BackColor = RGB(255, 0, 0)
        Case 20
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 156
            ColorSelect(2).Value = 156
            lblColor(20).BackColor = RGB(255, 156, 156)
        Case 21
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 210
            ColorSelect(2).Value = 210
            lblColor(21).BackColor = RGB(255, 210, 210)
        Case 22
            ColorSelect(0).Value = 255
            ColorSelect(1).Value = 235
            ColorSelect(2).Value = 235
            lblColor(22).BackColor = RGB(255, 235, 235)         ' Red
        '-------------------------------------------------------
        Case 23
            ColorSelect(0).Value = 0
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 0
            lblColor(23).BackColor = RGB(0, 255, 0)
        Case 24
            ColorSelect(0).Value = 156
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 156
            lblColor(24).BackColor = RGB(156, 255, 156)
        Case 25
            ColorSelect(0).Value = 210
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 210
            lblColor(25).BackColor = RGB(210, 255, 210)
        Case 26
            ColorSelect(0).Value = 235
            ColorSelect(1).Value = 255
            ColorSelect(2).Value = 235
            lblColor(26).BackColor = RGB(235, 255, 235)         ' Green
        '-------------------------------------------------------
        Case 27
            ColorSelect(0).Value = 250
            ColorSelect(1).Value = 250
            ColorSelect(2).Value = 250
            lblColor(27).BackColor = RGB(250, 250, 250)         ' Grey
        '=======================================================
        
    End Select
        
errhandler:
    Exit Sub
End Sub


Private Sub lblFinal_Click(Index As Integer)

    StatusBar1.Panels.Item(1).Text = lblFinal(Index).ToolTipText
    
End Sub

Private Sub lblFinal_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo errhandler
    
Dim N                   As Long
Dim P                   As Long
Dim Result              As Long
Dim msg                 As String
Dim tmp                 As String
Dim WebUsage            As String
Dim WebUsageBak         As String
Dim vRed                As Integer
Dim vGreen              As Integer
Dim vBlue               As Integer
Dim tRed                As Integer
Dim tGreen              As Integer
Dim tBlue               As Integer

Dim z()                 As String
Dim PosInRange          As Long
    
    ' set file_dirty flag
    File_Dirty_G = True
    
    ' remove Final color from display
    If Shift = 1 Then
        If MsgBox("Do you really wish to remove this color from the palette?     ", vbExclamation + vbOKCancel, "Remove Color", 0, 0) = 2 Then Exit Sub
        lblFinal(Index).BackColor = RGB(255, 255, 255)
        lblFinal(Index).Visible = False
        lblFinal(Index).ToolTipText = vbNullString
        txtFinal(Index).Visible = False
        txtFinal(Index + 10).Visible = False
        lblFinalNo(Index).Visible = False
        lblTarget(Index).Visible = False
        txtFinal(Index).Text = vbNullString
    
    'enter a description to selected color
    ElseIf Button = 2 Then
        WebUsageBak = Trim$(lblFinal(Index).ToolTipText)
        WebUsage = InputBox("Enter a short description of the web elements using the selected color in your website:", "Color Legend Editor", lblFinal(Index).ToolTipText)
                
        If Len(WebUsageBak) > 0 And Len(WebUsage) = 0 Then          ' when the user clicks 'Cancel'
            lblFinal(Index).ToolTipText = Space$(1) & WebUsageBak & Space$(1)
        Else
            lblFinal(Index).ToolTipText = Space$(1) & Trim$(WebUsage) & Space$(1)
        End If
        
    Else
        If Index = LastActiveIndex_G Then
            Exit Sub
        End If
        
        ' reset adjustment settings
        cmdReset_Click 1
        
        ' disable synchronising function
        chkLock(0).Value = 0
        chkLock(1).Value = 0
        chkLock(2).Value = 0
            
        ' split decimal string at ","'s
        z() = Split(txtFinal(Index + 10).Text, ",")
            
        ' convert scroll values to integers
        vRed = CInt(z(0))
        vGreen = CInt(z(1))
        vBlue = CInt(z(2))
            
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
        ColorSelect(0).Value = vRed
        ColorSelect(1).Value = vGreen
        ColorSelect(2).Value = vBlue
        
        ' color coding active color field
        For P = 0 To 10
            lblUsed(P).Visible = False
            txtColorCode(P).BackColor = &HE9F3F4
            lblRed(P).BackColor = &HE9F3F4
            lblGreen(P).BackColor = &HE9F3F4
            lblBlue(P).BackColor = &HE9F3F4
        Next P
        
        lblUsed(PosInRange).Visible = True
        txtColorCode(PosInRange).BackColor = &HFFFFFF
        lblRed(PosInRange).BackColor = &HFFFFFF
        lblGreen(PosInRange).BackColor = &HFFFFFF
        lblBlue(PosInRange).BackColor = &HFFFFFF
            
        LastActiveIndex_G = Index
    End If
    
    OverFlowWarning
    
errhandler:
    Exit Sub
End Sub

Private Sub lblLinkColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo errhandler
        
    Select Case txtLinkForeColorHex(Index).Index
    
        Case 1, 3, 5, 7, 9
            
            If Len(txtLinkForeColorHex(Index + 1).Text) = 0 Or _
               Len(txtLinkBackColorHex(Index + 1).Text) = 0 Then
                Exit Sub
            End If
            
            lblLinkColor(Index).ForeColor = "&H00" & InvertHexValue(txtLinkForeColorHex(Index + 1).Text)
            lblLinkColor(Index).BackColor = "&H00" & InvertHexValue(txtLinkBackColorHex(Index + 1).Text)
            
        Case 2, 4, 6, 8, 10
             
            If Len(txtLinkForeColorHex(Index - 1).Text) = 0 Or _
               Len(txtLinkBackColorHex(Index - 1).Text) = 0 Then
                Exit Sub
            End If
            
            lblLinkColor(Index).ForeColor = "&H00" & InvertHexValue(txtLinkForeColorHex(Index - 1).Text)
            lblLinkColor(Index).BackColor = "&H00" & InvertHexValue(txtLinkBackColorHex(Index - 1).Text)
            
    End Select
    
errhandler:
    Exit Sub
    
End Sub


Private Sub lblLinkColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo errhandler
        
    If Len(txtLinkForeColorHex(Index).Text) = 0 Or _
       Len(txtLinkBackColorHex(Index).Text) = 0 Then
       Exit Sub
    End If
    
    lblLinkColor(Index).ForeColor = "&H00" & InvertHexValue(txtLinkForeColorHex(Index).Text)
    lblLinkColor(Index).BackColor = "&H00" & InvertHexValue(txtLinkBackColorHex(Index).Text)
       
errhandler:
    Exit Sub
    
End Sub

Private Sub lblSample_Click(Index As Integer)
    
On Error GoTo errhandler
       
    ' display/hide small colored boxes denoting selected colors
    lblUsed(Index).BackColor = lblSample(Index).BackColor
    
    If lblUsed(Index).Visible = True Then
        lblUsed(Index).Visible = False
        txtColorCode(Index).BackColor = &HE9F3F4
        lblRed(Index).BackColor = &HE9F3F4
        lblGreen(Index).BackColor = &HE9F3F4
        lblBlue(Index).BackColor = &HE9F3F4
    Else
        lblUsed(Index).Visible = True
        txtColorCode(Index).BackColor = &HFFFFFF
        lblRed(Index).BackColor = &HFFFFFF
        lblGreen(Index).BackColor = &HFFFFFF
        lblBlue(Index).BackColor = &HFFFFFF
    End If
            
    'frmCMYK.Show
            
    'frmCMYK.txtColorValue.Text = lblSample(Index).BackColor
            
errhandler:
    Exit Sub
End Sub


Private Sub lblUsed_Click(Index As Integer)

On Error GoTo errhandler
    
    Me.Width = 9825
    ViewItem(0).Checked = True
    ViewItem(1).Checked = False
    
    Select Case SSTab1.Tab
        Case 0
            Select_General_Color Index
        Case 1
            Select_Link_Color Index
        Case 2
            Select_Text_Color Index
    End Select
    
errhandler:
    Exit Sub
End Sub

Private Sub lstWebElements_Click(Index As Integer)

On Error GoTo errhandler

Dim N                       As Integer
Dim P                       As Integer
Dim tmp                     As String
Dim msg1                    As String
Dim msg2                    As String
Dim txt1                    As String
Dim txt2                    As String
Dim txt3                    As String
    
    tmp = Right$(lstWebElements(Index).List(lstWebElements(Index).ListIndex), 2)
    If Len(tmp) > 0 Then
        N = CInt(tmp)
    Else
        Exit Sub
    End If
            
    If Index = 0 Then
        For P = 0 To lstWebElements(1).ListCount - 1
            If CInt(Right$(lstWebElements(1).List(P), 2)) = N Then
                lstWebElements(1).Selected(P) = True
            Else
                lstWebElements(1).Selected(P) = False
            End If
        Next P
    Else
        For P = 0 To lstWebElements(0).ListCount - 1
            If CInt(Right$(lstWebElements(0).List(P), 2)) = N Then
                lstWebElements(0).Selected(P) = True
            Else
                lstWebElements(0).Selected(P) = False
            End If
        Next P
    End If
    
    Call lblFinal_MouseDown(N, 1, 0, 0, 0)
    
    txt1 = Trim$(Left$(lstWebElements(1).List(lstWebElements(1).ListIndex), 6))
    txt2 = Left$(txt1, 2) & Space(1) & Mid$(txt1, 3, 2) & Space(1) & Right$(txt1, 2)
    txt3 = CInt("&H" & Left$(txt1, 2)) & ", " & CInt("&H" & Mid$(txt1, 3, 2)) & ", " & CInt("&H" & Right$(txt1, 2))
        
    Clipboard.Clear
    lblClipBoard(0).Caption = vbNullString
    lblClipBoard(1).Caption = vbNullString
    
    Clipboard.SetText txt1, 1
    lblClipBoard(0).Caption = txt2
    lblClipBoard(1).Caption = txt3
    
errhandler:
    Exit Sub
    
End Sub

Private Sub mAbout_Click()

On Error GoTo errhandler

    frmAbout.Show

errhandler:
    Exit Sub
End Sub

Private Sub mPick_Click()

On Error GoTo errhandler
    
    frmColor.Show
    
errhandler:
    Exit Sub
End Sub

Private Sub mPref_Click()

On Error GoTo errhandler
    
    frmPreferences.Show
    
errhandler:
    Exit Sub
End Sub


Private Sub RecentFileItem_Click(Index As Integer)
    
On Error GoTo errhandler

    If Index = 0 Then
        FileOpen GetOpenFilePath
        
    Else
        FileOpen RecentFileItem(Index).Caption
        
    End If
    
errhandler:
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

On Error GoTo errhandler
    
Dim N                       As Long

    IndexFore_M = -1
    IndexBack_M = -1
    
    Select Case SSTab1.Tab
        Case 0
            StatusBar1.Panels.Item(2).Text = " Color palette for Page Elements"
            
        Case 1
            StatusBar1.Panels.Item(2).Text = " Color palette for Links (Normal & Hover)"
            
        Case 2
            StatusBar1.Panels.Item(2).Text = " Color palette for Text (Forecolor & Backcolor)"
            
        Case 3
            lstWebElements(0).Clear
            lstWebElements(1).Clear
            For N = 1 To 10
                If Len(lblFinal(N).ToolTipText) > 0 Then
                    lstWebElements(0).AddItem Trim$(lblFinal(N).ToolTipText) & " = " & Trim$(txtFinal(N).Text) & Space$(100) & Format$(N, "00")
                    lstWebElements(1).AddItem Trim$(txtFinal(N).Text) & " = " & Trim$(lblFinal(N).ToolTipText) & Space$(100) & Format$(N, "00")
                End If
            Next N
            
    End Select
    
errhandler:
    Exit Sub
End Sub
Private Sub txtBackColor_Click(Index As Integer)

On Error GoTo errhandler

    IndexFore_M = -1
    IndexBack_M = txtBackColor(Index).Index
    
errhandler:
    Exit Sub
End Sub


Private Sub txtBackColorHex_Click(Index As Integer)

On Error GoTo errhandler

    IndexFore_M = -1
    IndexBack_M = txtBackColorHex(Index).Index
    
errhandler:
    Exit Sub
End Sub


Private Sub txtColorCode_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo errhandler

Dim cCode               As String
    
    ' exit if selected field index is greater that 0
    If Index > 0 Then
        Exit Sub
    End If
        
    ' convert to upper case and remove #
    cCode = Trim$(UCase$(Replace(txtColorCode(0).Text, "#", vbNullString)))
    
    ' implement manually entered R,G,B hex values
    ColorSelect(0).Value = CInt("&H" & Left$(cCode, 2))
    ColorSelect(1).Value = CInt("&H" & Mid$(cCode, 3, 2))
    ColorSelect(2).Value = CInt("&H" & Right$(cCode, 2))
            
errhandler:
    Exit Sub
End Sub


Private Sub txtColorCode_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo errhandler
    
    ' update display
    ColorSelect_Scroll 0
    
errhandler:
    Exit Sub
    
End Sub

Private Sub txtColorCode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo errhandler

Dim txt1 As String
Dim txt2 As String
Dim txt3 As String

    txt1 = Trim$(txtColorCode(Index).Text)
    txt2 = Left$(txt1, 2) & Space(1) & Mid$(txt1, 3, 2) & Space(1) & Right$(txt1, 2)
    txt3 = CInt("&H" & Left$(txt1, 2)) & ", " & CInt("&H" & Mid$(txt1, 3, 2)) & ", " & CInt("&H" & Right$(txt1, 2))
        
    Clipboard.Clear
    lblClipBoard(0).Caption = vbNullString
    lblClipBoard(1).Caption = vbNullString
    
    Clipboard.SetText txt1, 1
    lblClipBoard(0).Caption = txt2
    lblClipBoard(1).Caption = txt3
    
errhandler:
    Exit Sub
End Sub

Private Sub txtFinal_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo errhandler

Dim txt1 As String
Dim txt2 As String
Dim txt3 As String

    txt1 = Trim$(txtFinal(Index).Text)
    txt2 = Left$(txt1, 2) & Space(1) & Mid$(txt1, 3, 2) & Space(1) & Right$(txt1, 2)
    txt3 = CInt("&H" & Left$(txt1, 2)) & ", " & CInt("&H" & Mid$(txt1, 3, 2)) & ", " & CInt("&H" & Right$(txt1, 2))
    
    Clipboard.Clear
    lblClipBoard(0).Caption = vbNullString
    lblClipBoard(1).Caption = vbNullString
    
    Clipboard.SetText txt1, 1
                              
    lblClipBoard(0).Caption = txt2
    lblClipBoard(1).Caption = txt3
    
errhandler:
    Exit Sub
End Sub


Private Sub txtForeColor_Click(Index As Integer)

On Error GoTo errhandler

    IndexFore_M = txtForeColor(Index).Index
    IndexBack_M = -1
    
errhandler:
    Exit Sub
End Sub


Private Sub txtForeColorHex_Click(Index As Integer)

On Error GoTo errhandler

    IndexFore_M = txtForeColorHex(Index).Index
    IndexBack_M = -1
    
errhandler:
    Exit Sub
End Sub


Private Sub txtLinkBackColor_Click(Index As Integer)

On Error GoTo errhandler

    IndexFore_M = -1
    IndexBack_M = txtLinkBackColor(Index).Index
    
errhandler:
    Exit Sub
End Sub


Private Sub txtLinkBackColorHex_Click(Index As Integer)

On Error GoTo errhandler

    IndexFore_M = -1
    IndexBack_M = txtLinkBackColorHex(Index).Index
    
errhandler:
    Exit Sub
End Sub


Private Sub txtLinkForeColor_Click(Index As Integer)

On Error GoTo errhandler

    IndexFore_M = txtLinkForeColor(Index).Index
    IndexBack_M = -1
    
errhandler:
    Exit Sub
End Sub


Private Sub txtLinkForeColorHex_Click(Index As Integer)

On Error GoTo errhandler

    IndexFore_M = txtLinkForeColorHex(Index).Index
    IndexBack_M = -1
    
errhandler:
    Exit Sub
End Sub


Private Sub txtNewDec_Click(Index As Integer)

On Error GoTo errhandler
    
    txtNewDec(Index).SelStart = 0
    txtNewDec(Index).SelLength = 3
    DoEvents
    
errhandler:
    Exit Sub
End Sub

Private Sub txtNewDec_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo errhandler
    
    ' convert to upper case
    If InStr("1234567890", Chr(KeyAscii)) = 0 Then KeyAscii = 0
        
errhandler:
    Exit Sub
End Sub


Private Sub txtNewDec_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo errhandler

Dim hexRed              As String
Dim hexGreen            As String
Dim hexBlue             As String
    
    ' correct if entered R,G,B values excede 255
    If CInt(txtNewDec(0).Text) > 255 Then txtNewDec(0).Text = 255
    If CInt(txtNewDec(1).Text) > 255 Then txtNewDec(1).Text = 255
    If CInt(txtNewDec(2).Text) > 255 Then txtNewDec(2).Text = 255
    
    ' implement entered R,G,B values
    ColorSelect(0).Value = CInt(txtNewDec(0).Text)
    ColorSelect(1).Value = CInt(txtNewDec(1).Text)
    ColorSelect(2).Value = CInt(txtNewDec(2).Text)
    
    ' convert R,G,B values to hex and compose string
    hexRed = Hex(ColorSelect(0).Value): If Len(hexRed) = 1 Then hexRed = "0" & hexRed
    hexGreen = Hex(ColorSelect(1).Value): If Len(hexGreen) = 1 Then hexGreen = "0" & hexGreen
    hexBlue = Hex(ColorSelect(2).Value): If Len(hexBlue) = 1 Then hexBlue = "0" & hexBlue
    
    ' add # in front and display hex value
    txtNewHexValue.Text = Space(1) & hexRed & hexGreen & hexBlue
    
    ' update display
    ColorSelect_Scroll 0
    
errhandler:
    Exit Sub
End Sub

Private Sub txtNewHexValue_Click()

On Error GoTo errhandler
    
    txtNewHexValue.SelStart = 0
    txtNewHexValue.SelLength = 7
        
errhandler:
    Exit Sub
End Sub

Private Sub txtNewHexValue_KeyPress(KeyAscii As Integer)
    
On Error GoTo errhandler
    
    ' convert to upper case
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    
    ' prevent illegal characters in hex color code
    If InStr("ABCDEF1234567890", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
errhandler:
    Exit Sub
End Sub

Public Sub txtNewHexValue_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo errhandler

Dim cCode               As String
Dim cPos                As Long
    
    ' convert to upper case and remove #
    cCode = Trim$(UCase$(Replace(txtNewHexValue.Text, "#", vbNullString)))
    
    'split hex R,G,B value into separete decimal values for scroll boxes
    ColorSelect(0).Value = CInt("&H" & Left$(cCode, 2))
    ColorSelect(1).Value = CInt("&H" & Mid$(cCode, 3, 2))
    ColorSelect(2).Value = CInt("&H" & Right$(cCode, 2))
    
    ' display hex values for R,G,B
    txtNewDec(0).Text = Format$(ColorSelect(0).Value, "000")
    txtNewDec(1).Text = Format$(ColorSelect(1).Value, "000")
    txtNewDec(2).Text = Format$(ColorSelect(2).Value, "000")
    
    ' store position of cursor while updating
    cPos = txtNewHexValue.SelStart
        
    ' update display
    ColorSelect_Scroll 0
            
    ' insert R,G,B hex code in text field
    txtNewHexValue.Text = Space(1) & cCode
    
    ' return cursor position
    txtNewHexValue.SelStart = cPos + 1
    
errhandler:
    Exit Sub
End Sub


Private Sub ViewItem_Click(Index As Integer)
    
On Error GoTo errhandler
    
    SSTab1.SetFocus
    DoEvents
    
    Select Case Index
        Case 0
            StatusBar1.style = 0
            Me.Width = 9825
            ViewItem(0).Checked = True
            ViewItem(1).Checked = False
        Case 1
            StatusBar1.style = 1
            Me.Width = 5285
            ViewItem(0).Checked = False
            ViewItem(1).Checked = True
    End Select
    
errhandler:
    Exit Sub
End Sub


