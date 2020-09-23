Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String)
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cY As Long, ByVal wFlags As Long)

Public int_Gradient_Interval    As Integer
Public int_Gradient_Factor      As Integer
Public sng_Gradient_Factor      As Single

Public LastActiveIndex_G        As Long

Public dir_Default              As String
Public dir_OpenFileTitle        As String
Public dir_SaveFileTitle        As String
Public dir_OpenFolder           As String
Public dir_SaveFolder           As String
Public pref_UserName            As String
Public pref_Description         As String
Public pref_Initials            As String
Public File_Dirty_G             As Boolean

Public x()                      As String
Public y()                      As String

Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const swDot = "."

Public Type FILETIME
        dwLowDateTime           As Long
        dwHighDateTime          As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes             As Long
   ftCreationTime               As FILETIME
   ftLastAccessTime             As FILETIME
   ftLastWriteTime              As FILETIME
   nFileSizeHigh                As Long
   nFileSizeLow                 As Long
   dwReserved0                  As Long
   dwReserved1                  As Long
   cFileName                    As String * MAX_PATH
   cAlternate                   As String * 14
End Type

Public Function FormatHexCodeInv(ByVal vHex As String) As String
    
On Error GoTo errhandler
        
    ' adjust string length to 6 characters by adding "0"'s in front
    vHex = String$(6 - Len(vHex), "0") & vHex
    
    ' invert RGB sequence
    FormatHexCodeInv = Right$(vHex, 2) & Mid$(vHex, 3, 2) & Left$(vHex, 2)

errhandler:
    Exit Function
End Function

Public Sub form_StayOnTop( _
                           ByVal frmOnTop As Object, _
                           ByVal hwnd As String, _
                           ByVal Position As String _
                         )

On Error GoTo errhandler

Dim TpPT        As Long
Dim TpPL        As Long
Dim TpPH        As Long
Dim TpPW        As Long
Dim h           As Long
Dim oT          As Long
Dim Success     As Boolean
            
    Select Case UCase$(hwnd)
        Case "ABSOLUTE":        oT = -1   ' HWND_TOPMOST      above all non-topmost windows, maintains its topmost position when deactivated
        Case "TOPMOST":         oT = 2    ' HWND_NOTOPMOST    above all non-topmost windows
        Case "TOP":             oT = 0    ' HWND_TOP          top of z-order
        Case "BOTTOM":          oT = 1    ' HWND_BOTTOM       bottom of z-order
        Case Else:              oT = 0    ' HWND              top of z-order
    End Select
    
h = frmOnTop.Height
       
    '--------------------------------------------------------------------------
    ' Get conversion factors from twips to pixel for X and Y
    '--------------------------------------------------------------------------
    TpPH = Screen.TwipsPerPixelY    ' = pixels
    TpPW = Screen.TwipsPerPixelX

    Select Case UCase$(Position)
        Case " "                                                ' No change
            TpPT = frmOnTop.Top
            TpPL = frmOnTop.Left
            
            ' prevent form position from being outside screen area
            If frmOnTop.Top > Screen.Height - 100 Or _
               frmOnTop.Left > Screen.Width - 100 Then
                TpPT = 100
                TpPL = 100
            End If
            
        Case "TL", "LT"                                             ' Top Left
            TpPT = 0 / TpPH
            TpPL = 0 / TpPW
        
        Case "TR", "RT"                                             ' Top Right
            TpPT = 0
            TpPL = (Screen.Width - frmOnTop.Width) / TpPW
            
        Case "BL", "LB"                                             ' Botom Left
            TpPT = (Screen.Height - frmOnTop.Height) / TpPH
            TpPL = 0 / TpPW
            
        Case "BR", "RB"                                             ' Bottom Right
            TpPT = (Screen.Height - frmOnTop.Height) / TpPH
            TpPL = (Screen.Width - frmOnTop.Width) / TpPW
            
        Case "C"                                                    ' Center
            TpPT = (Screen.Height / 2 - frmOnTop.Height / 2) / TpPH
            TpPL = (Screen.Width / 2 - frmOnTop.Width / 2) / TpPW
            
        Case Else                                                   ' No change
            TpPT = frmOnTop.Top
            TpPL = frmOnTop.Left
            
            ' prevent form position from being outside screen area
            If frmOnTop.Top > Screen.Height - 100 Or _
               frmOnTop.Left > Screen.Width - 100 Then
                TpPT = 100
                TpPL = 100
            End If
    End Select
    
    TpPH = frmOnTop.Height / TpPH   ' FormHeight in pixels
    TpPW = frmOnTop.Width / TpPW    ' FormWidth in pixels
            
    '----------------------------------------------------------------------
    ' Call API SetWindowPos sub
    '----------------------------------------------------------------------
    SetWindowPos frmOnTop.hwnd, oT, TpPL, TpPT, TpPW, TpPH, 0 '&H40
        
    frmOnTop.Height = h
    
    Exit Sub
    
errhandler:
    Exit Sub
            
End Sub


            
Public Sub ColorListFromTemplate(ByVal Template As Object)

On Error GoTo errhandler

Dim N                   As Long
Dim P                   As Long
Dim tmp                 As String
Dim count               As Long
Dim vHex                As String
Dim vRed                As String
Dim vGreen              As String
Dim vBlue               As String
    
    ReDim X1(1 To 32, 1 To 2) As String
    ReDim Y1(1 To 32, 1 To 2) As String
    
    For N = 1 To 10
        frmMain.lblTarget(N).Visible = False
    Next N
    
    count = 0
    
    ' get unique list of BACKCOLORS on template (without white and black)
    For N = 0 To 12 ' (Element 0 - 12)
        For P = 1 To 32
            If X1(P, 1) = Template.Element(N).BackColor Or _
                          Template.Element(N).BackColor = CLng(&H0) Or _
                          Template.Element(N).BackColor < 0 Or _
                          Template.Element(N).BackColor = CLng(&HFFFFFF) Then
                GoTo NextN1
            End If
        Next P
        count = count + 1
        X1(count, 1) = Template.Element(N).BackColor
        X1(count, 2) = " Element"
NextN1:
    Next N
    
    ' get unique list of FORECOLORS on template (without white and black)
    For N = 0 To 12 ' (Element 0 - 12)
        For P = 1 To 32
            If X1(P, 1) = Template.Element(N).ForeColor Or _
                         Template.Element(N).ForeColor = CLng(&H0) Or _
                         Template.Element(N).ForeColor < 0 Or _
                         Template.Element(N).ForeColor = CLng(&HFFFFFF) Then
                GoTo NextN2
            End If
        Next P
        count = count + 1
        X1(count, 1) = Template.Element(N).ForeColor
        X1(count, 2) = " Text"
NextN2:
    Next N
    
    ' get unique list of LINE BORDERCOLORS on template (without white and black)
    For N = 1 To 8 ' (hLine 1 - 8)
        For P = 1 To 32
            If X1(P, 1) = Template.hLine(N).BorderColor Or _
                         Template.hLine(N).BorderColor = CLng(&H0) Or _
                         Template.hLine(N).BorderColor < 0 Or _
                         Template.hLine(N).BorderColor = CLng(&HFFFFFF) Then
                GoTo NextN3
            End If
        Next P
        count = count + 1
        X1(count, 1) = Template.hLine(N).BorderColor
        X1(count, 2) = " Line"
NextN3:
    Next N

    ' transfer to Y() removing empty entries
    count = 0
    For N = 1 To 32
        If Len(X1(N, 1)) > 0 Then
            count = count + 1
            Y1(count, 1) = X1(N, 1)
            Y1(count, 2) = X1(N, 2)
        End If
    Next N
    
    ' update Final list om frmMain
    count = 0
    For N = 1 To 32
        
        ' lists ALL colors used in template on frmMain display starting from
        ' position 1. - rest of the Final colors are left unaffected
        If Len(Y1(N, 1)) > 0 Then
            count = count + 1
            frmMain.lblFinal(count).Visible = True
            frmMain.lblFinalNo(count).Visible = True
            frmMain.txtFinal(count).Visible = True
            frmMain.txtFinal(count + 10).Visible = True
            frmMain.lblTarget(count).Visible = True
            
            frmMain.lblFinal(count).BackColor = CLng(Y1(N, 1))
            
            ' VB - for unknown reasons - lists hex color values as BGR (not RGB!)
            ' hence the hex value must be reversed
            vHex = Hex(CLng(Y1(count, 1)))
            vHex = String$(6 - Len(vHex), "0") & vHex
            vHex = Right$(vHex, 2) & Mid$(vHex, 3, 2) & Left$(vHex, 2)
            
            ' get decimal values for RGB from hex value
            vRed = Format$(CLng("&H" & Left$(vHex, 2)), "000")
            vGreen = Format$(CLng("&H" & Mid$(vHex, 3, 2)), "000")
            vBlue = Format$(CLng("&H" & Right$(vHex, 2)), "000")
            
            frmMain.lblFinalNo(count).Caption = N
            frmMain.txtFinal(count).Text = " " & vHex
            frmMain.txtFinal(count + 10).Text = " " & vRed & ", " & vGreen & ", " & vBlue
            frmMain.lblTarget(count).Caption = Y1(count, 2)
        End If
        
    Next N

errhandler:
    Exit Sub
End Sub




Public Function GetDirectoryPath(ByVal NewPath As String) As String

On Error GoTo errhandler
     
Dim pos                 As Long
    
    If Len(NewPath) > 0 Then
        pos = InStrRev(NewPath, "\")
        If pos > 0 And pos < Len(NewPath) Then
            NewPath = Left$(NewPath, pos)
        End If
    Else
        GoTo errhandler
    End If
            
    GetDirectoryPath = NewPath
    
    Exit Function
    
errhandler:
    GetDirectoryPath = dir_Default
    Exit Function
    
End Function

Public Function swFolderExists(ByVal Path As String) As Boolean
    
On Error GoTo errhandler
    
    If Len(Path) = 0 Then GoTo errhandler
    
Dim att                 As Long
Dim swArchieve          As Boolean
Dim swFolder            As Boolean
Dim swVolume            As Boolean
Dim swSystem            As Boolean
Dim swHidden            As Boolean
Dim swReadOnly          As Boolean
Dim swNormal            As Boolean
Dim hFile               As Long
Dim WFD                 As WIN32_FIND_DATA
        
    ' Path is a FOLDER:                                 swFolderExists = TRUE if folder exists
    '                                                   swFolderExists = FALSE if folder does not exist
    If Right$(Path, 1) = "\" Then
        Path = Left$(Path, Len(Path) - 1)
        hFile = FindFirstFile(Path, WFD)
        swFolderExists = (hFile <> INVALID_HANDLE_VALUE) And (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
        Call FindClose(hFile)
        Exit Function
    
    ' Path is a FOLDER or a FILE
    Else
    
        ' no WILDCARD characters in Path
        If InStr(Path, "*") = 0 Then
            hFile = FindFirstFile(Path, WFD)
            swFolderExists = (hFile <> INVALID_HANDLE_VALUE) And (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
            Call FindClose(hFile)
            If swFolderExists Then
                Exit Function                           ' swFolderExists = TRUE - Path is a FOLDER and exists                       swFolderExists = TRUE if Path is a folder and exists
            Else
                att = GetFileAttributes(Path)
                If att = -1 Then
                    swFolderExists = False                       ' swFolderExists = FALSE - Path does not exist
                    Exit Function
                Else
                    If att / 32 >= 1 Then swArchieve = True: att = att Mod 32
                    If att / 16 >= 1 Then swFolder = True: att = att Mod 16
                    If att / 8 >= 1 Then swVolume = True: att = att Mod 8
                    If att / 4 >= 1 Then swVolume = True: att = att Mod 4
                    If att / 2 >= 1 Then swHidden = True: att = att Mod 2
                    If att = 1 Then swReadOnly = True Else swNormal = True
                    If swNormal Or swReadOnly Then
                        swFolderExists = True                    ' swFolderExists = TRUE - Path is a FILE and exits
                        Exit Function
                    Else
                        swFolderExists = False                   ' swFolderExists = FALSE - Path is a file and does not exist
                        Exit Function
                    End If
                End If
            End If
            
        ' WILDCARD characters in Path
        Else
            If InStr(GetFileTitleFromPath(Path), "*") > 0 Then
                If Len(Dir(Path)) > 0 Then
                    swFolderExists = True                        ' swFolderExists = TRUE - FOLDER exists with at least one FILE
                    Exit Function
                Else
                    swFolderExists = False                       ' swFolderExists = FALSE - FOLDER is empty
                    Exit Function
                End If
            End If
        End If
    End If
        
errhandler:
    swFolderExists = False
    Exit Function
        
End Function

Public Function GetFileTitleFromPath(ByVal FullPath As String) As String

On Error GoTo errhandler
     
Dim pos                 As Long
Dim tmp                 As String
            
    ' exit with nothing if NewPath ends with a backslash
    If Right$(FullPath, 1) = "\" Or Len(FullPath) < 4 Then
        GoTo errhandler
    End If
        
    ' get file title
    pos = InStrRev(FullPath, "\")
    If pos > 0 Then
        tmp = Mid$(FullPath, pos + 1)
    Else
        GoTo errhandler
    End If
            
    GetFileTitleFromPath = UCase$(Trim$(tmp))
    
    Exit Function
        
errhandler:
    GetFileTitleFromPath = vbNullString
    Exit Function
    
End Function

Public Function swFileExists(ByVal FilePath As String) As Boolean
        
On Error GoTo ErrorHandler

Dim hFile               As Long
Dim WFD                 As WIN32_FIND_DATA

    hFile = FindFirstFile(FilePath, WFD)
    swFileExists = hFile <> INVALID_HANDLE_VALUE
   
    Call FindClose(hFile)
        
    Exit Function
    
ErrorHandler:
    swFileExists = True
    Exit Function

End Function






