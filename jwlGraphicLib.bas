Attribute VB_Name = "jwlGraphicLib"
'Graphics Lib v. 0.1.1
'Functions and procedures for working with graphics.
'Created: July 16 2003 - JWL

'Modified: July 23 2003 Added RGB - JWL
'Modified: July 28 2003 Added Web Hex - JWL
'Modified: August 1 2003 Added HLS - JWL
'Modified: August 6 2003 Debuged HLS - JWL
'Modified: August 9 2003 Added CMYK - JWL
'Modified: August 10 2003 Added lighten & darken - JWL
'Modified: August 11 2003 Added invert - JWL
'Modified: July 23 2004 Added Blend Alpha - JWL


'University of Illinois/NCSA Open Source License

'Copyright (c)  2003 Joseph W. Lumbley
'All rights reserved.
'Developed by: Open VB Group - http://jlumbley.tripod.com

'Permission is hereby granted, free of charge, to any person obtaining a
'copy of this software and associated documentation files (the "Software"),
'to deal with the Software without restriction, including without limitation
'the rights to use, copy, modify, merge, publish, distribute, sublicense,
'and/or sell copies of the Software, and to permit persons to whom the
'Software is furnished to do so, subject to the following conditions:

'Redistributions of source code must retain the above copyright notice, this
'list of conditions and the following disclaimers.

'Redistributions in binary form must reproduce the above copyright notice,
'this list of conditions and the following disclaimers in the documentation
'and/or other materials provided with the distribution.

'Neither the name of Open VB Group nor the names of its contributors may
'be used to endorse or promote products derived from this Software
'without specific prior written permission.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY
'KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
'WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'CONTRIBUTORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
'DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
'TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH
'THE SOFTWARE OR THE USE OR OTHER DEALINGS WITH THE SOFTWARE.


'This is part of an open source visual basic project called
'The Open VB Graphics Editor developed by the Open VB Group.
'Project Maintainer: Joseph W. Lumbley

'For project information, updates, or to contribute please visit our website.
'To report bugs or suggest improvements please visit our website.
'Open VB Group website - http://jlumbley.tripod.com

'Open VB Group is a trademark of Joseph W. Lumbley.


'OSI Certified Open Source Software.
'http://opensource.org

'Algorithm References:

'The RGB algorithms are derived from the MSDN Library Visual Studio 6
'article called "Determining RGB Color Values"

'The Web HEX algorithms are derived from the MSDN Library VS 6
'article called "Color Table"

'The HLS algorithms are derived from the MSDN Library VS 6
'Article ID: Q29240. This article cites as a reference: Foley and Van Dam,
'"Fundamentals of Interactive Computer Graphics," Pages 618-19.
'The HSL algorithms are also derived from code by Andrew Gray
'Link: http://abstractvb.com/code.asp?F=50&P=1&A=927

'The CMYK algorithms are derived from code by Saifudheen A.A.
'keraleeyan@msn.com, www.saifu.5u.com
'You can view the code, comment on the code/and or vote on it at:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=38924&lngWId=1

'**************************************************************************************
' algorithm from: http://www.fsref.com/Fatal/FE070401.SHTML
'RGB TO CMYK - EasyRBG Algorithm
'
'     C ' = 1 - (R/range)      C = (C' - K') / (1 - K')
'     M ' = 1 - (G/range)      M = (M' - K') / (1 - K')
'     y ' = 1 - (B/range)      Y = (Y' - K') / (1 - K')
'     K ' = MIN(C',Y',M')      K = K'
'
'     CMYK values are in 0 - 1 range, multiply by range to convert to same range as RGB.
'     C 'Y'M' are CYM values, in 0 - 1 range, multiply by range to convert to same range as RGB.
'     Use this formula to convert CMY to CMYK.
'
'CMYK TO RGB - EasyRBG Algorithm
'
'     C ' = (C*(1-K)+K)        R = (1 - C') * range
'     M ' = (M*(1-K)+K)        G = (1 - M') * range
'     y ' = (Y*(1-K)+K)        B = (1 - Y') * range
'
'     CMYK values are in 0 - 1 range, divide by range to convert from same range as RGB.
'     C 'Y'M' are CYM values, in 0 - 1 range, multiply by range to convert to same range as RGB.
'     Use this formula to convert CMYK to CMY.
'
'NOTE:
'     MAX( ) - maximum of values in parenthesis.
'     MIN( ) - miniumum of values in parenthesis.
'     range  - RGB range, usually 255 (FFh) or 100 (%)
'**************************************************************************************
Option Explicit

Public Enum jwlFillStyle
    Solid = 0
    Transparent = 1
    Horizontal_Lines = 2
    Vertical_Lines = 3
    Upward_Diagonal = 4
    Downward_Diagonal = 5
    Cross = 6
    Diagonal_Cross = 7
End Enum

Public Const HueMAX = 239, SatMAX = 240, LumMAX = 240

Public Sub jwlRec(ByRef DrawTo As Control, ByVal Top As Integer, ByVal Left As Integer, _
ByVal Width As Integer, ByVal Height As Integer, _
Optional ByVal BorderWidth As Integer = -1, _
Optional ByVal BorderColor As Long = -1, _
Optional ByVal FillStyle As jwlFillStyle = -1, _
Optional ByVal FillColor As Long = -1)

'A simple rectangle procedure to replace the line box step procedure

    Select Case FillStyle
        Case 1, -1
            If BorderWidth <> -1 Then
                DrawTo.DrawWidth = BorderWidth
            Else
                DrawTo.DrawWidth = 1
            End If
            
            If BorderColor <> -1 Then
                DrawTo.Line (Left, Top)-Step(Width, Height), BorderColor, B
            Else
                DrawTo.Line (Left, Top)-Step(Width, Height), , B
            End If
            
        Case 0, 2 - 7
            If FillColor <> -1 Then
                DrawTo.Line (Left, Top)-Step(Width, Height), FillColor, BF
            Else
                DrawTo.Line (Left, Top)-Step(Width, Height), , BF
            End If
            
    End Select
    
End Sub

Public Function jwlRed(ByVal LongClrValue As Long) As Integer

    'Input
    'A long color.
    
    'Returns
    'The red value of the RGB color of the long color.
    
    If LongClrValue > 255 Then
        jwlRed = LongClrValue Mod 256
    Else
        jwlRed = LongClrValue
    End If
    
End Function

Public Function jwlGreen(ByVal LongClrValue As Long) As Integer

    'Input
    'A long color.
    
    'Returns
    'The green value of the RGB color of the long color.

    Dim NewLong As Long
    
    If LongClrValue > 65535 Then
        NewLong = LongClrValue Mod 65536
    Else
        NewLong = LongClrValue
    End If
    
    If LongClrValue > 255 Then
        jwlGreen = Int(NewLong / 256)
    Else
        jwlGreen = 0
    End If
    
End Function

Public Function jwlBlue(ByVal LongClrValue As Long) As Integer

    'Input
    'A long color.
    
    'Returns
    'The blue value of the RGB color of the long color.

    If LongClrValue > 65535 Then
        jwlBlue = Int(LongClrValue / 65536)
    Else
        jwlBlue = 0
    End If
    
End Function

Public Function jwlWebHex(ByVal LongClrValue As Long) As String

    'Input
    'A long color.
    
    'Returns
    'The web hex color of the long color.
    
    Dim r, g, b As Integer
    Dim Rhex, Ghex, Bhex As String
    
    r = jwlRed(LongClrValue)
    g = jwlGreen(LongClrValue)
    b = jwlBlue(LongClrValue)
    
    Rhex = Right("0" & Hex(r), 2)
    Ghex = Right("0" & Hex(g), 2)
    Bhex = Right("0" & Hex(b), 2)
    
    jwlWebHex = "#" & Rhex & Ghex & Bhex
    
End Function

Private Function jwlMaxRGB(ByVal LongClrValue As Long) As Integer

    Dim r, g, b As Integer

    r = jwlRed(LongClrValue)
    g = jwlGreen(LongClrValue)
    b = jwlBlue(LongClrValue)
    
    If r > g Then
        If r > b Then
            jwlMaxRGB = r
        Else
            jwlMaxRGB = b
        End If
    Else
        If g > b Then
            jwlMaxRGB = g
        Else
            jwlMaxRGB = b
        End If
    End If
    
End Function

Private Function jwlMinRGB(ByVal LongClrValue As Long) As Integer

    Dim r, g, b As Integer

    r = jwlRed(LongClrValue)
    g = jwlGreen(LongClrValue)
    b = jwlBlue(LongClrValue)
    
    If r < g Then
        If r < b Then
            jwlMinRGB = r
        Else
            jwlMinRGB = b
        End If
    Else
        If g < b Then
            jwlMinRGB = g
        Else
            jwlMinRGB = b
        End If
    End If
    
End Function

Public Function jwlHue(ByVal LongClrValue As Long) As Integer

    'Input
    'A long color.
    
    'Returns
    'The hue value of the HSL color of the long color.
    

    'Debugged to match the output of the color dialog box

    Dim r, g, b, MaxRGB, MinRGB As Integer
    Dim h, RMN, GMN, BMN As Double

    r = jwlRed(LongClrValue)
    g = jwlGreen(LongClrValue)
    b = jwlBlue(LongClrValue)
    
    MaxRGB = jwlMaxRGB(LongClrValue)
    MinRGB = jwlMinRGB(LongClrValue)
    
    If MaxRGB = MinRGB Then
        h = 160 'Debug: the color dialog box uses hue 160 for grayscale
    Else
        RMN = (((MaxRGB - r) * (239 / 6)) + 0.5) / (MaxRGB - MinRGB)
        GMN = (((MaxRGB - g) * (239 / 6)) + 0.5) / (MaxRGB - MinRGB)
        BMN = (((MaxRGB - b) * (239 / 6)) + 0.5) / (MaxRGB - MinRGB)
        Select Case MaxRGB
            Case r
                h = BMN - GMN
            Case g
                h = (239 / 3) + RMN - BMN
            Case b
                h = ((2 * 239) / 3) + GMN - RMN
        End Select
        If h < 0 Then h = h + 239
    End If
    
    jwlHue = CInt(h)
    
End Function

Public Function jwlSaturation(ByVal LongClrValue As Long) As Integer

    'Input
    'A long color.
    
    'Returns
    'The saturation value of the HSL color of the long color.

    Dim MaxRGB, MinRGB As Integer
    Dim S, L As Double

    MaxRGB = jwlMaxRGB(LongClrValue)
    MinRGB = jwlMinRGB(LongClrValue)
    
    L = (((MaxRGB + MinRGB) * 240) + 255) / (2 * 255)
    
    If MaxRGB = MinRGB Then
        S = 0
    Else
        If L <= (240 / 2) Then
            S = (((MaxRGB - MinRGB) * 240) + 0.5) / (MaxRGB + MinRGB)
        Else
            S = (((MaxRGB - MinRGB) * 240) + 0.5) / (2 * 255 - (MaxRGB + MinRGB))
        End If
    End If
    
    jwlSaturation = CInt(S)
    
End Function

Public Function jwlLuminescence(ByVal LongClrValue As Long) As Integer
    
    'Input
    'A long color.
    
    'Returns
    'The luminescence value of the HSL color of the long color.
    
    
    'Debugged to match the output of the color dialog box
    Dim MaxRGB, MinRGB As Integer
    Dim L As Double
    
    MaxRGB = jwlMaxRGB(LongClrValue)
    MinRGB = jwlMinRGB(LongClrValue)
    
    L = (((MaxRGB + MinRGB) * 240) + 255) / (2 * 255)
    
    jwlLuminescence = Int(L) 'Debug: rounding error fixed to match the color dialog box
    
    
End Function

Public Function jwlCyan(ByVal LongClrValue As Long) As Integer
    
    'Input
    'A long color.
    
    'Returns
    'The cyan value of the CMYK color of the long color.

    Dim r, g, b, C, M, y, K, MinColor As Integer

    r = jwlRed(LongClrValue)
    g = jwlGreen(LongClrValue)
    b = jwlBlue(LongClrValue)
    
    C = 255 - r
    M = 255 - g
    y = 255 - b
    
    K = IIf(C < M, C, M)
    
    If y < K Then
        K = y
    End If
        
    If K > 0 Then
        C = C - K
        M = M - K
        y = y - K
    End If
        
    MinColor = IIf(C < M, C, M)
    MinColor = IIf(y < MinColor, y, MinColor)
    MinColor = IIf((MinColor + K) > 255, 255 - K, MinColor)
    
    If C - MinColor <> 0 Then
        jwlCyan = 100 * 255 / (C - MinColor)
    Else
        jwlCyan = 100
    End If

End Function

Public Function jwlMagenta(ByVal LongClrValue As Long) As Integer

    'Input
    'A long color.
    
    'Returns
    'The magenta value of the CMYK color of the long color.

    Dim r, g, b, C, M, y, K, MinColor As Integer

    r = jwlRed(LongClrValue)
    g = jwlGreen(LongClrValue)
    b = jwlBlue(LongClrValue)
    
    C = 255 - r
    M = 255 - g
    y = 255 - b
    
    K = IIf(C < M, C, M)
    
    If y < K Then
        K = y
    End If
        
    If K > 0 Then
        C = C - K
        M = M - K
        y = y - K
    End If
        
    MinColor = IIf(C < M, C, M)
    MinColor = IIf(y < MinColor, y, MinColor)
    MinColor = IIf((MinColor + K) > 255, 255 - K, MinColor)
    
    If M - MinColor <> 0 Then
        jwlMagenta = 100 * (M - MinColor) / 255
    Else
        jwlMagenta = 100
    End If
        
End Function

Public Function jwlYellow(ByVal LongClrValue As Long) As Integer

    'Input
    'A long color.
    
    'Returns
    'The yellow value of the CMYK color of the long color.

    Dim r, g, b, C, M, y, K, MinColor As Integer

    r = jwlRed(LongClrValue)
    g = jwlGreen(LongClrValue)
    b = jwlBlue(LongClrValue)
    
    C = 255 - r
    M = 255 - g
    y = 255 - b
    
    K = IIf(C < M, C, M)
    
    If y < K Then
        K = y
    End If
        
    If K > 0 Then
        C = C - K
        M = M - K
        y = y - K
    End If
        
    MinColor = IIf(C < M, C, M)
    MinColor = IIf(y < MinColor, y, MinColor)
    MinColor = IIf((MinColor + K) > 255, 255 - K, MinColor)
    
    If y - MinColor <> 0 Then
        jwlYellow = 100 * (y - MinColor) / 255
    Else
        jwlYellow = 100
    End If
        
End Function

Public Function jwlBlack(ByVal LongClrValue As Long) As Integer

    'Input
    'A long color.
    
    'Returns
    'The black value of the CMYK color of the long color.

    Dim r, g, b, C, M, y, K, MinColor As Integer

    r = jwlRed(LongClrValue)
    g = jwlGreen(LongClrValue)
    b = jwlBlue(LongClrValue)
    
    C = 255 - r
    M = 255 - g
    y = 255 - b
    
    K = IIf(C < M, C, M)
    
    If y < K Then
        K = y
    End If
        
    If K > 0 Then
        C = C - K
        M = M - K
        y = y - K
    End If
        
    MinColor = IIf(C < M, C, M)
    MinColor = IIf(y < MinColor, y, MinColor)
    MinColor = IIf((MinColor + K) > 255, 255 - K, MinColor)
    
    If K + MinColor <> 0 Then
        jwlBlack = 100 * (K + MinColor) / 255
    Else
        jwlBlack = 100
    End If
        
End Function

Public Function jwlDarken(ByVal LongClrValue As Long, ByVal DarkenValue As Integer) As Long

    'Input
    'A color and darken value.
    
    'Returns
    'A darken color.
    
    'Valid range for darken value is 0 to 240.

    Dim h, S, L As Integer
    
    h = jwlHue(LongClrValue)
    S = jwlSaturation(LongClrValue)
    L = jwlLuminescence(LongClrValue)
    
    '0 is black 240 is white
    If (L - DarkenValue) >= 0 Then
        L = L - DarkenValue
    Else
        L = 0
    End If
    
    'Convert HSL darken color to long color.
    'Return long color.
    jwlDarken = jwlHSL(h, S, L)
    
End Function

Public Function jwlLighten(ByVal LongClrValue As Long, ByVal LightenValue As Integer) As Long
    
    'Input
    'A color and lighten value.
    
    'Returns
    'A lighten color.
    
    'Valid range for lighten value is 0 to 240.
    
    Dim h, S, L As Integer
    
    'Get HSL for long color.
    h = jwlHue(LongClrValue)
    S = jwlSaturation(LongClrValue)
    L = jwlLuminescence(LongClrValue)
    
    'Find lighten color.
    '0 is black 240 is white
    If (L + LightenValue) <= 240 Then
        L = L + LightenValue
    Else
        L = 240
    End If
    
    'Convert HSL lighten color to long lighten color.
    'Return long lighten color.
    jwlLighten = jwlHSL(h, S, L)
    
End Function

Public Function jwlInvert(ByVal LongClrValue As Long) As Long

    Dim r, g, b As Integer
    
    'Get RGB for long color.
    r = jwlRed(LongClrValue)
    g = jwlGreen(LongClrValue)
    b = jwlBlue(LongClrValue)
    
    'Find inverted color.
    r = 255 - r
    g = 255 - g
    b = 255 - b
    
    'Convert RGB inverted color to long inverted color.
    'Return long inverted color
    jwlInvert = RGB(r, g, b)
        
End Function

Public Function jwlHSL(ByVal Hue As Integer, ByVal Saturation As Integer, ByVal Luminance As Integer) As Long

    'Converts HSL color to long color.
    'Input
    'HSL color.
    'Returns
    'The long color of the HSL color.

    Dim pHue, pSat, pLum, pRed, pGreen, pBlue, temp2, temp1 As Single
    Dim temp3() As Single
    Dim r, g, b, N As Integer
    
    ReDim temp3(0 To 2)
    
    pHue = Hue / 239 '239
    pSat = Saturation / 240 '239
    pLum = Luminance / 240 '239
    
    If pSat = 0 Then
        pRed = pLum
        pGreen = pLum
        pBlue = pLum
    Else
        If pLum < 0.5 Then
            temp2 = pLum * (1 + pSat)
        Else
            temp2 = pLum + pSat - pLum * pSat
        End If
        
        temp1 = 2 * pLum - temp2
        
        temp3(0) = pHue + 1 / 3
        temp3(1) = pHue
        temp3(2) = pHue - 1 / 3
        
        For N = 0 To 2
            If temp3(N) < 0 Then temp3(N) = temp3(N) + 1
            If temp3(N) > 1 Then temp3(N) = temp3(N) - 1
            
            If 6 * temp3(N) < 1 Then
                temp3(N) = temp1 + (temp2 - temp1) * 6 * temp3(N)
            Else
                If 2 * temp3(N) < 1 Then
                    temp3(N) = temp2
                Else
                    If 3 * temp3(N%) < 2 Then
                        temp3(N%) = temp1 + (temp2 - temp1) * ((2 / 3) - temp3(N%)) * 6
                    Else
                        temp3(N%) = temp1
                    End If
                End If
            End If
        Next N%
    
        pRed = temp3(0)
        pGreen = temp3(1)
        pBlue = temp3(2)
    End If

    r = Int(pRed * 255)
    g = Int(pGreen * 255)
    b = Int(pBlue * 255)
    
    If r < 0 Then r = 0
    If g < 0 Then g = 0
    If b < 0 Then b = 0
    
    jwlHSL = RGB(r, g, b)
    
End Function

Public Function jwlBlendAlpha(ByVal OriginColor As Long, _
ByVal DestinationColor As Long, ByVal AlphaValue As Integer) As Long

    'Input
    'Two colors and alpha value.
    
    'Returns
    'A blend of the two colors.
    
    'Valid range for alpha value is 1 to 255.

    Dim OriginR, OriginG, OriginB As Integer
    Dim DestinR, DestinG, DestinB As Integer
    Dim RedSpan, GreenSpan, BlueSpan As Single
    Dim RedMN, GreenMN, BlueMN As Single
    Dim ResultRed, ResultGreen, ResultBlue As Integer
    
    'Get RGB for origin color.
    OriginR = jwlRed(OriginColor)
    OriginG = jwlGreen(OriginColor)
    OriginB = jwlBlue(OriginColor)
    
    'Get RGB for destination color.
    DestinR = jwlRed(DestinationColor)
    DestinG = jwlGreen(DestinationColor)
    DestinB = jwlBlue(DestinationColor)
    
    'Find Red blend color.
    RedSpan = DestinR - OriginR
    RedMN = RedSpan / 255
    ResultRed = CInt((AlphaValue * RedMN) + OriginR)
    
    'Find Green blend color.
    GreenSpan = DestinG - OriginG
    GreenMN = GreenSpan / 255
    ResultGreen = CInt((AlphaValue * GreenMN) + OriginG)
    
    'Find Blue blend color.
    BlueSpan = DestinB - OriginB
    BlueMN = BlueSpan / 255
    ResultBlue = CInt((AlphaValue * BlueMN) + OriginB)
    
    'Convert RGB blend color to long blend color.
    'Return long blend color
    jwlBlendAlpha = RGB(ResultRed, ResultGreen, ResultBlue)
    
End Function
