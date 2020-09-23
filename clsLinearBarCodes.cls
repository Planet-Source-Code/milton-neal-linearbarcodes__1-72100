VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinearBarCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' *************************************************************************
' Milton Neal, Perth Western Australia. miltonneal@arach.net.au. Jan. 2009
' *************************************************************************
Public Enum eSymbology
    Code39 = 1
    Code39_Extd = 2
    Code128_Auto = 3
    Code128_A = 4
    Code128_B = 5
    Code128_C = 6
    Code2of5 = 7
    CodeI2of5 = 8
    Code93 = 9
    Codabar = 10
    MSI = 11
    Code11 = 12
End Enum

Public Enum eBarTextPlacement
    TP_UNDER = 0
    TP_ABOVE = 1
End Enum

Public Enum eBarTextAlign
    ALIGN_CENTRE = 0
    ALIGN_LEFT = 1
    ALIGN_RIGHT = 2
    ALIGN_FULL = 3
End Enum

Private BC39_CharSet As String * 43
Private BC39_Code(43) As String

Private BC93_CharSet As String * 47
Private BC93_Code(47) As String

Private BC128A_CharSet As String * 95
Private BC128B_CharSet As String * 95
Private BC128C_CharSet As String * 10
Private BC128_Code(106) As String

Private Codabar_CharSet As String * 16
Private Codabar_Code(19)

Private MSI_Charset As String * 10
Private MSI_Code(11)

Private BC11_Charset As String * 11
Private BC11_Code(11)

Private BC2of5_Charset As String * 10
Private BC2of5_Code(11)
Private BCI2of5_code(11)

Private cHDC As Long
Private cSymbology As eSymbology    'Barcode symbology to use
Private cBarX As Single             'Width of the bar in printer pixels
Private cBarMultiplier As Integer   'Bar width multiplier
Private cBarRatio As Single         'Ratio between narrow and wide bars for code 39
Private cBarColor As OLE_COLOR      'Color to draw the barcode
Private cBarCodeData As String      'Barcode data
Private cBarRotation As Integer     'Direction to print/display barcode
Private cBarCodeTop As Long         'Top start position in pixels
Private cBarCodeLeft As Long        'Left start position in pixel
Private cBarCodeHeight As Long      'Height of the bars (not including text)in pixels
Private cBarTextGap As Long         'Gap between the barcode and the HR text in pixels

Private cHumanReadable As Boolean               'Human readable or barcode only
Private cBarTextFont As UDT_BarTextFont         'Holds the font information
Private cBarTextPosition As eBarTextPlacement   'Placement of the HR text, above or below barcode
Private cBarTextAlign As eBarTextAlign          'Alignment of the text within the barcode
Private cDisplayStartStop As Boolean            'True if to show "*" in the code 39/93 HR text
Private cUseCheckDigit As Boolean               'True if to use check digit if the barcode has the option
Private cDisplayCheckDigit As Boolean           'True if to show the check digit in the code 39 HR text

Dim BarStartX As Long           'X point to start barcode
Dim BarStartY As Long           'Y point to start barcode
Dim NarrowBar As Single         'Narrow bar
Dim WideBar As Single           'Wide bar
Dim BarCodeWidth As Long        'Width of barcode in pixels

Dim HRBarText As String         'Holds the printable HR barcode text
Dim TextStartX As Long          'X point for text
Dim TextStartY As Long          'Y point for text
Dim HRTextHeight As Long        'HR text height in pixels
Dim HRTextWidth As Long         'HR text width in Pixels
Dim EncodedData As String       'Holds the encoded data for creating barcode
Dim BarBit As String * 1        'Holds the "width" value of a bar

Private cErrorNumber As Long
Private cErrDescription As String
Const B_ErrNoData As Long = 5940
Const S_ErrNoData As String = "No barcode data specified."
Const B_ErrSymbology As Long = 5950
Const S_ErrSymbology As String = "Invalid or no symbology selected."
Const B_ErrDraw As Long = 5960
Const S_ErrDraw As String = "Error generating barcode."
Const B_ErrInvalidChar As Long = 5970
Const S_ErrInvalidChar As String = "Invalid character in the Barcode."
Const B_ErrInvalidRatio As Long = 5980
Const S_ErrInvalidRatio As String = "Invalid barcode ratio."
Const B_ErrInvalidControl As Long = 5990
Const S_ErrInvalidControl As String = "Invalid control to output the barcode."

'===============================================
'PROPERTIES
'===============================================

Public Property Let BarcodeOutput(dest As Object)
    On Error GoTo err
    cHDC = dest.hdc
    On Error GoTo 0
    Exit Property
err:
    cErrorNumber = B_ErrInvalidControl
    cErrDescription = S_ErrInvalidControl
    err.Raise cErrorNumber, "clsBarCode.BarcodeOutput", cErrDescription
End Property

Public Property Get Symbology() As eSymbology
    Symbology = cSymbology
End Property
Public Property Let Symbology(bSym As eSymbology)
    cSymbology = bSym
End Property

Public Property Get BarXFactor() As Single
    BarXFactor = cBarX
End Property
Public Property Let BarXFactor(xbWidth As Single)
    cBarX = xbWidth
End Property


Public Property Get BarMultiplier() As Integer
    BarMultiplier = cBarMultiplier
End Property
Public Property Let BarMultiplier(bMulti As Integer)
    If bMulti > 10 Then bMulti = 10
    If bMulti < 1 Then bMulti = 1
    cBarMultiplier = bMulti
End Property

Public Property Get BarRatio() As Single
    BarRatio = cBarRatio * 2
End Property
Public Property Let BarRatio(bRatio As Single)
    Select Case cSymbology
        Case Code39, Code39_Extd, CodeI2of5
            If bRatio < 2 Or bRatio > 3 Then
                cErrorNumber = B_ErrInvalidRatio
                cErrDescription = S_ErrInvalidRatio
                err.Raise cErrorNumber, "clsBarCode.BarRatio", cErrDescription
                Exit Property
            End If
            cBarRatio = bRatio / 2
        Case Else
            cBarRatio = 1
    End Select
End Property

Public Property Get UseCheckDigit() As Boolean
    UseCheckDigit = cUseCheckDigit
End Property
Public Property Let UseCheckDigit(cDigit As Boolean)
    cUseCheckDigit = cDigit
End Property

Public Property Get BarRotation() As Integer
    BarRotation = cBarRotation
End Property
Public Property Let BarRotation(bRotation As Integer)
    If bRotation <> 0 And bRotation <> 90 _
    And bRotation <> 180 And bRotation <> 270 Then
        cBarRotation = 0
    Else: cBarRotation = bRotation
    End If
End Property

Public Property Get BarCodeData() As String
    BarCodeData = cBarCodeData
End Property

Public Property Let BarCodeData(bcData As String)
    Dim iCntr As Integer
    Dim sDataChar As String
    Dim iAscValue As Integer
    
    bcData = Trim(bcData)
    'Check each barcode character for being valid
    Select Case cSymbology
        Case Code39
            cBarCodeData = bcData
            For iCntr = 1 To Len(bcData)
                sDataChar = Mid(bcData, iCntr, 1)
                If InStr(1, BC39_CharSet, sDataChar) = 0 Then GoTo Raise_Error
            Next iCntr
            Encode39 bcData
            
        Case Code39_Extd
            For iCntr = 1 To Len(bcData)
                iAscValue = AscW(Mid(bcData, iCntr, 1))
                If iAscValue < 0 Or iAscValue > 127 Then GoTo Raise_Error
            Next iCntr
            Encode39Extd (bcData)
            
        Case Code128_Auto
            For iCntr = 1 To Len(bcData)
                iAscValue = AscW(Mid(bcData, iCntr, 1))
                If iAscValue < 0 Or iAscValue > 127 Then GoTo Raise_Error
            Next iCntr
            Encode128 bcData
            
        Case Code128_A
            For iCntr = 1 To Len(bcData)
                sDataChar = Mid(bcData, iCntr, 1)
                If InStr(1, BC128A_CharSet, sDataChar) = 0 Then GoTo Raise_Error
            Next iCntr
            Encode128a bcData
            
        Case Code128_B
            For iCntr = 1 To Len(bcData)
            sDataChar = Mid(bcData, iCntr, 1)
            If InStr(1, BC128B_CharSet, sDataChar) = 0 Then GoTo Raise_Error
            Next iCntr
            Encode128b bcData
            
        Case Code128_C
            For iCntr = 1 To Len(bcData)
            sDataChar = Mid(bcData, iCntr, 1)
            If InStr(1, BC128C_CharSet, sDataChar) = 0 Then GoTo Raise_Error
            Next iCntr
            Encode128c bcData
            
        Case Code2of5, CodeI2of5
            For iCntr = 1 To Len(bcData)
                sDataChar = Mid(bcData, iCntr, 1)
                If InStr(1, BC2of5_Charset, sDataChar) = 0 Then GoTo Raise_Error
            Next iCntr
            If cSymbology = Code2of5 Then Encode2of5 bcData Else EncodeI2of5 bcData
            
        Case Code93
            For iCntr = 1 To Len(bcData)
                sDataChar = Mid(bcData, iCntr, 1)
                If InStr(1, BC93_CharSet, sDataChar) = 0 Then GoTo Raise_Error
            Next iCntr
            Encode93 bcData
            
        Case Codabar
            For iCntr = 1 To Len(bcData)
                sDataChar = Mid(bcData, iCntr, 1)
                If InStr(1, Codabar_CharSet, sDataChar) = 0 Then GoTo Raise_Error
            Next iCntr
            EncodeCodaBar bcData
            
        Case MSI
            For iCntr = 1 To Len(bcData)
                sDataChar = Mid(bcData, iCntr, 1)
                If InStr(1, MSI_Charset, sDataChar) = 0 Then GoTo Raise_Error
            Next iCntr
            EncodeMSI bcData
            
        Case Code11
            For iCntr = 1 To Len(bcData)
                sDataChar = Mid(bcData, iCntr, 1)
                If InStr(1, BC11_Charset, sDataChar) = 0 Then GoTo Raise_Error
            Next iCntr
            Encode11 bcData
    End Select
    cBarCodeData = bcData
    Exit Property
    
Raise_Error:
    cErrorNumber = B_ErrInvalidChar
    cErrDescription = S_ErrInvalidChar
    err.Raise cErrorNumber, "clsBarCode.BarCodeData", cErrDescription
End Property

Public Property Get BarColor() As OLE_COLOR
    BarColor = cBarColor
End Property

Public Property Let BarColor(bColor As OLE_COLOR)
    cBarColor = bColor
End Property

Public Property Get BarCodeX() As Long
    BarCodeX = cBarCodeLeft
End Property

Public Property Let BarCodeX(X As Long)
    cBarCodeLeft = X
End Property

Public Property Get BarCodeY() As Long
    BarCodeY = cBarCodeTop
End Property

Public Property Let BarCodeY(Y As Long)
    cBarCodeTop = Y
End Property

Public Property Get BarcodeHeight() As Long
    BarcodeHeight = cBarCodeHeight
End Property

Public Property Let BarcodeHeight(bHeight As Long)
    If bHeight < 8 Then bHeight = 8
    cBarCodeHeight = bHeight
End Property

Public Property Get BarTextGap() As Long
    BarTextGap = cBarTextGap
End Property

Public Property Let BarTextGap(bGap As Long)
    cBarTextGap = bGap
End Property

Public Property Get HRText() As Boolean
    HRText = cHumanReadable
End Property

Public Property Let HRText(bReadable As Boolean)
    cHumanReadable = bReadable
End Property

Public Property Get HRTextPlacement() As eBarTextPlacement
    HRTextPlacement = cBarTextPosition
End Property

Public Property Let HRTextPlacement(bPlacement As eBarTextPlacement)
    cBarTextPosition = bPlacement
End Property

Public Property Get HRTextAlignment() As eBarTextAlign
    HRTextAlignment = cBarTextAlign
End Property

Public Property Let HRTextAlignment(bTextAlign As eBarTextAlign)
    cBarTextAlign = bTextAlign
End Property

Public Property Get DisplayStopStart() As Boolean
    DisplayStopStart = cDisplayStartStop
End Property
Public Property Let DisplayStopStart(sSSChar As Boolean)
    cDisplayStartStop = sSSChar
End Property

Public Property Get DisplayCheckDigit() As Boolean
    DisplayCheckDigit = cDisplayCheckDigit
End Property
Public Property Let DisplayCheckDigit(sChkDigit As Boolean)
    cDisplayCheckDigit = sChkDigit
End Property

Public Property Get TotalBarWidth() As Long
    If cBarCodeData = "" Then
        cErrorNumber = B_ErrNoData
        cErrDescription = S_ErrNoData
        err.Raise cErrorNumber, "clsBarCode.BarWidth", cErrDescription
        Exit Property
    End If
    Select Case cBarRotation
        Case 0, 180
            TotalBarWidth = BarCodeWidth
        Case 90, 270
            If cHumanReadable Then
                Call GetTextSize(cHDC, cBarTextFont, HRBarText, HRTextWidth, HRTextHeight)
                TotalBarWidth = cBarCodeHeight + HRTextHeight + cBarTextGap
            Else: TotalBarWidth = cBarCodeHeight
            End If
    End Select
End Property

Public Property Get TotalBarHeight() As Long
    If cBarCodeData = "" Then
        cErrorNumber = B_ErrNoData
        cErrDescription = S_ErrNoData
        err.Raise cErrorNumber, "clsBarCode.BarWidth", cErrDescription
        Exit Property
    End If
    Select Case cBarRotation
        Case 0, 180
            If cHumanReadable Then
                Call GetTextSize(cHDC, cBarTextFont, HRBarText, HRTextWidth, HRTextHeight)
                TotalBarHeight = cBarCodeHeight + HRTextHeight + cBarTextGap
            Else: TotalBarHeight = cBarCodeHeight
            End If
        Case 90, 270
            TotalBarHeight = BarCodeWidth
    End Select
End Property

Public Sub BarTextFont(bFont As String, bFontSize As Single, bFontBold As Boolean, bFontItalic As Boolean)
    If bFont = "" Then
        bFont = cBarTextFont.FontName
        bFontSize = cBarTextFont.FontSize
        bFontBold = cBarTextFont.FontBold
        bFontItalic = cBarTextFont.FontItalic
    Else
        cBarTextFont.FontName = bFont
        cBarTextFont.FontSize = bFontSize
        cBarTextFont.FontBold = bFontBold
        cBarTextFont.FontItalic = bFontItalic
    End If
End Sub

Public Function DrawBarCode()

    'Check that a symbology has been set
    If cSymbology < 1 Or cSymbology > 12 Then
        cErrorNumber = B_ErrSymbology
        cErrDescription = S_ErrSymbology
        err.Raise cErrorNumber, "clsBarCode.Draw_BarCode", cErrDescription
        Exit Function
    End If
    'Check that the barcode data has been set
    If cBarCodeData = "" Then
        cErrorNumber = B_ErrNoData
        cErrDescription = S_ErrNoData
        err.Raise cErrorNumber, "clsBarCode.BarWidth", cErrDescription
        Exit Function
    End If
    
    Select Case cBarRotation
        Case 0
            Draw_Normal
        Case 90
            Draw_90
        Case 180
            Draw_180
        Case 270
            Draw_270
    End Select
Exit Function

Err_Handler:
    cErrorNumber = B_ErrInvalidControl
    cErrDescription = S_ErrInvalidControl
    err.Raise cErrorNumber, "clsBarCode39.Draw_BarCode", cErrDescription
End Function

Private Sub Draw_Normal()

    Dim NextBar As Single   'Postion to start the next bar
    Dim iCntr As Integer    'Loop counter
    Dim iSeg As Single      'For printing ALIGN_FULL text
    Dim bColor As Long      'Current bar color
    Dim rtn As Long
    
    'Set the starting points for the barcode and the text
    rtn = GetTextSize(cHDC, cBarTextFont, HRBarText, HRTextWidth, HRTextHeight)   'Text size
    If rtn = 0 Then GoTo Raise_Error

    BarStartX = cBarCodeLeft: BarStartY = cBarCodeTop
    If cBarTextPosition = TP_UNDER Then
        TextStartY = cBarCodeTop + cBarCodeHeight + cBarTextGap
    Else
        TextStartY = cBarCodeTop
        BarStartY = cBarCodeTop + HRTextHeight + cBarTextGap
    End If
    
    'Draw the barcode
    bColor = vbWhite
    NextBar = BarStartX
    For iCntr = 1 To Len(EncodedData)
        'Toggle the color
        If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
        BarBit = Mid(EncodedData, iCntr, 1)
        Select Case BarBit
            'Narrow Bar
            Case "1"
                If bColor = vbWhite Then
                    NextBar = NextBar + NarrowBar
                Else
                    rtn = DrawBar(cHDC, CLng(NextBar), BarStartY, CLng(NextBar + NarrowBar), BarStartY + cBarCodeHeight, bColor)
                    If rtn = 0 Then GoTo Raise_Error
                    NextBar = NextBar + NarrowBar
                End If
            'Wide Bar
            Case "2", "3", "4"
                WideBar = NarrowBar * BarBit * cBarRatio
                If bColor = vbWhite Then
                    NextBar = NextBar + WideBar
                Else
                    rtn = DrawBar(cHDC, CLng(NextBar), BarStartY, CLng(NextBar + WideBar), BarStartY + cBarCodeHeight, bColor)
                    If rtn = 0 Then GoTo Raise_Error
                    NextBar = NextBar + WideBar
                End If
        End Select
    Next iCntr
    
    'Draw the barcode text
    If cHumanReadable Then
        Select Case cBarTextAlign
            'Centred
            Case ALIGN_CENTRE
                TextStartX = BarStartX + (BarCodeWidth - HRTextWidth) \ 2
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Left
            Case ALIGN_LEFT
                TextStartX = BarStartX
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Right
            Case ALIGN_RIGHT
                TextStartX = BarStartX + (BarCodeWidth - HRTextWidth)
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Print each character spaced equally across the width of the barcode
            Case ALIGN_FULL
                iSeg = BarCodeWidth / Len(HRBarText)
                For iCntr = 1 To Len(HRBarText)
                    'Get the width of each character
                    Call GetTextSize(cHDC, cBarTextFont, Mid(HRBarText, iCntr, 1), HRTextWidth, HRTextHeight)
                    TextStartX = BarStartX + (iCntr * iSeg) - (iSeg \ 2) - (HRTextWidth \ 2)
                    rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, Mid(HRBarText, iCntr, 1))
                    If rtn = 0 Then GoTo Raise_Error
                Next iCntr
        End Select
    End If
Exit Sub
    
Raise_Error:
    cErrorNumber = B_ErrDraw
    cErrDescription = S_ErrDraw
    err.Raise cErrorNumber, "clsBarCode39.Draw_Normal", cErrDescription
End Sub

Private Sub Draw_90()

    Dim NextBar As Single   'Postion to start the next bar
    Dim iCntr As Integer    'Loop counter
    Dim iSeg As Single      'For printing ALIGN_FULL text
    Dim bColor As Long      'Current bar color
    Dim rtn As Long
    
    'Set the starting points for the barcode and the text
    rtn = GetTextSize(cHDC, cBarTextFont, HRBarText, HRTextWidth, HRTextHeight)
    If rtn = 0 Then GoTo Raise_Error
    
    BarStartX = cBarCodeLeft: BarStartY = cBarCodeTop + BarCodeWidth
    If cBarTextPosition = TP_UNDER Then
        TextStartX = cBarCodeLeft + cBarCodeHeight + cBarTextGap
    Else
        TextStartX = cBarCodeLeft
        BarStartX = cBarCodeLeft + HRTextHeight + cBarTextGap
    End If
    
    'Draw the barcode
    bColor = vbWhite
    NextBar = BarStartY
    For iCntr = 1 To Len(EncodedData)
        'Toggle the color
        If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
        BarBit = Mid(EncodedData, iCntr, 1)
        Select Case BarBit
            'Narrow Bar
            Case "1"
                If bColor = vbWhite Then
                    NextBar = NextBar - NarrowBar
                Else
                    rtn = DrawBar(cHDC, BarStartX, CLng(NextBar - NarrowBar), BarStartX + cBarCodeHeight, CLng(NextBar), bColor)
                    If rtn = 0 Then GoTo Raise_Error
                    NextBar = NextBar - NarrowBar
                End If
            'Wide Bar
            Case "2", "3", "4"
                WideBar = NarrowBar * BarBit * cBarRatio
                If bColor = vbWhite Then
                    NextBar = NextBar - WideBar
                Else
                    rtn = DrawBar(cHDC, BarStartX, CLng(NextBar - WideBar), BarStartX + cBarCodeHeight, CLng(NextBar), bColor)
                    If rtn = 0 Then GoTo Raise_Error
                    NextBar = NextBar - WideBar
                End If
            End Select
    Next iCntr
    
    'Draw the barcode text
    If cHumanReadable Then
        Select Case cBarTextAlign
            'Centred
            Case ALIGN_CENTRE
                TextStartY = BarStartY - ((BarCodeWidth - HRTextWidth) \ 2)
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Left
            Case ALIGN_LEFT
                TextStartY = BarStartY
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Right
            Case ALIGN_RIGHT
                TextStartY = BarStartY - (BarCodeWidth - HRTextWidth)
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Print each character spaced equally across the width of the barcode
            Case ALIGN_FULL
                iSeg = BarCodeWidth / Len(HRBarText)
                For iCntr = 1 To Len(HRBarText)
                    'Get the width of each character
                    Call GetTextSize(cHDC, cBarTextFont, Mid(HRBarText, iCntr, 1), HRTextWidth, HRTextHeight)
                    TextStartY = BarStartY - (iCntr * iSeg) + (iSeg \ 2) + (HRTextWidth \ 2)
                    rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, Mid(HRBarText, iCntr, 1))
                    If rtn = 0 Then GoTo Raise_Error
                Next iCntr
        End Select
    End If

Exit Sub

Raise_Error:
    cErrorNumber = B_ErrDraw
    cErrDescription = S_ErrDraw
    err.Raise cErrorNumber, "clsBarCode39.Draw_90", cErrDescription
End Sub

Private Sub Draw_180()

    Dim NextBar As Single   'Postion to start the next bar
    Dim iCntr As Integer    'Loop counter
    Dim iSeg As Single      'For printing ALIGN_FULL text
    Dim bColor As Long      'Current bar color
    Dim rtn As Long
    
    'Set the starting points for the barcode and the text
    rtn = GetTextSize(cHDC, cBarTextFont, HRBarText, HRTextWidth, HRTextHeight)
    If rtn = 0 Then GoTo Raise_Error
    
    BarStartX = cBarCodeLeft + BarCodeWidth: BarStartY = cBarCodeTop
    If cBarTextPosition = TP_UNDER Then
        TextStartY = cBarCodeTop + HRTextHeight
        BarStartY = cBarCodeTop + HRTextHeight + cBarTextGap
    Else
        TextStartY = cBarCodeTop + cBarCodeHeight + cBarTextGap + HRTextHeight
        BarStartY = cBarCodeTop
    End If

    'Draw the barcode
    bColor = vbWhite
    NextBar = BarStartX
    For iCntr = 1 To Len(EncodedData)
        'Toggle the color
        If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
        BarBit = Mid(EncodedData, iCntr, 1)
        Select Case BarBit
            'Narrow Bar
            Case "1"
                If bColor = vbWhite Then
                    NextBar = NextBar - NarrowBar
                Else
                    rtn = DrawBar(cHDC, CLng(NextBar - NarrowBar), BarStartY, CLng(NextBar), BarStartY + cBarCodeHeight, bColor)
                    If rtn = 0 Then GoTo Raise_Error
                    NextBar = NextBar - NarrowBar
                End If
            'Wide Bar
            Case "2", "3", "4"
                WideBar = NarrowBar * BarBit * cBarRatio
                If bColor = vbWhite Then
                    NextBar = NextBar - WideBar
                Else
                    rtn = DrawBar(cHDC, CLng(NextBar - WideBar), BarStartY, CLng(NextBar), BarStartY + cBarCodeHeight, bColor)
                    If rtn = 0 Then GoTo Raise_Error
                    NextBar = NextBar - WideBar
                End If
            End Select
    Next iCntr
    
    'Print the barcode text
    If cHumanReadable Then
        Select Case cBarTextAlign
            'Centred
            Case ALIGN_CENTRE
                TextStartX = BarStartX - ((BarCodeWidth - HRTextWidth) \ 2)
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Left
            Case ALIGN_LEFT
                TextStartX = BarStartX
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Right
            Case ALIGN_RIGHT
                TextStartX = BarStartX - (BarCodeWidth - HRTextWidth)
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Print each character spaced equally across the width of the barcode
            Case ALIGN_FULL
                iSeg = BarCodeWidth / Len(HRBarText)
                For iCntr = 1 To Len(HRBarText)
                    'Get the width of each character
                    Call GetTextSize(cHDC, cBarTextFont, Mid(HRBarText, iCntr, 1), HRTextWidth, HRTextHeight)
                    TextStartX = BarStartX - (iCntr * iSeg) + (iSeg \ 2) + (HRTextWidth \ 2)
                    rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, Mid(HRBarText, iCntr, 1))
                    If rtn = 0 Then GoTo Raise_Error
                Next iCntr
        End Select
    End If
Exit Sub

Raise_Error:
    cErrorNumber = B_ErrDraw
    cErrDescription = S_ErrDraw
    err.Raise cErrorNumber, "clsBarCode39.Draw_180", cErrDescription
End Sub

Private Sub Draw_270()
    Dim NextBar As Single   'Postion to start the next bar
    Dim iCntr As Integer    'Loop counter
    Dim iSeg As Single      'For printing ALIGN_FULL text
    Dim bColor As Long      'Current bar color
    Dim rtn As Long
    
    'Set the starting points for the barcode and the text
    rtn = GetTextSize(cHDC, cBarTextFont, HRBarText, HRTextWidth, HRTextHeight)
    If rtn = 0 Then GoTo Raise_Error
    
    BarStartX = cBarCodeLeft: BarStartY = cBarCodeTop
    If cBarTextPosition = TP_UNDER Then
        TextStartX = cBarCodeLeft + HRTextHeight
        BarStartX = cBarCodeLeft + HRTextHeight + cBarTextGap
    Else
        TextStartX = cBarCodeLeft + cBarCodeHeight + cBarTextGap + HRTextHeight
    End If
      
    'Draw the barcode
    bColor = vbWhite
    NextBar = BarStartY
    For iCntr = 1 To Len(EncodedData)
        'Toggle the color
        If bColor = vbWhite Then bColor = cBarColor Else bColor = vbWhite
        BarBit = Mid(EncodedData, iCntr, 1)
        Select Case BarBit
            'Narrow Bar
            Case "1"
                If bColor = vbWhite Then
                    NextBar = NextBar + NarrowBar
                Else
                    rtn = DrawBar(cHDC, BarStartX, CLng(NextBar), BarStartX + cBarCodeHeight, CLng(NextBar + NarrowBar), bColor)
                    If rtn = 0 Then GoTo Raise_Error
                    NextBar = NextBar + NarrowBar
                End If
            'Wide Bar
            Case "2", "3", "4"
                WideBar = NarrowBar * BarBit * cBarRatio
                If bColor = vbWhite Then
                    NextBar = NextBar + WideBar
                Else
                    rtn = DrawBar(cHDC, BarStartX, CLng(NextBar), BarStartX + cBarCodeHeight, CLng(NextBar + WideBar), bColor)
                    If rtn = 0 Then GoTo Raise_Error
                    NextBar = NextBar + WideBar
                End If
            End Select
    Next iCntr
    
    'Print the barcode text
    If cHumanReadable Then
        Select Case cBarTextAlign
            'Centred
            Case ALIGN_CENTRE
                TextStartY = BarStartY + (BarCodeWidth - HRTextWidth) \ 2
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Left
            Case ALIGN_LEFT
                TextStartY = BarStartY
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Right
            Case ALIGN_RIGHT
                TextStartY = BarStartY + (BarCodeWidth - HRTextWidth)
                rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, HRBarText)
                If rtn = 0 Then GoTo Raise_Error
            'Print each character spaced equally across the width of the barcode
            Case ALIGN_FULL
                iSeg = BarCodeWidth / Len(HRBarText)
                For iCntr = 1 To Len(HRBarText)
                    'Get the width of each character
                    Call GetTextSize(cHDC, cBarTextFont, Mid(HRBarText, iCntr, 1), HRTextWidth, HRTextHeight)
                    TextStartY = BarStartY + (iCntr * iSeg) - (iSeg \ 2) - (HRTextWidth \ 2)
                    rtn = DrawBarText(cHDC, TextStartX, TextStartY, cBarTextFont, cBarColor, cBarRotation, Mid(HRBarText, iCntr, 1))
                    If rtn = 0 Then GoTo Raise_Error
                Next iCntr
        End Select
    End If
Exit Sub

Raise_Error:
    cErrorNumber = B_ErrDraw
    cErrDescription = S_ErrDraw
    err.Raise cErrorNumber, "clsBarCode39.Draw_270", cErrDescription
End Sub

'=====================================
'CODE 39 FUNCTIONS
'=====================================

Private Function Encode39(DataToEncode As String)
    Dim iAscVal As Integer
    Dim iCntr As Integer
    Dim TotalWieght As Integer
    Dim iWeight As Integer
    Dim sDataChar As String * 1
    Dim sChkDigit As String * 1
    
    HRBarText = DataToEncode
    If cUseCheckDigit Then
        'Calculate the check character
        For iCntr = 1 To Len(DataToEncode)
            iAscVal = AscW(Mid(DataToEncode, iCntr, 1))
            sDataChar = (Mid(DataToEncode, iCntr, 1))
            TotalWieght = TotalWieght + (InStr(1, BC39_CharSet, sDataChar)) - 1
        Next iCntr
        iWeight = TotalWieght Mod 43
        'Get back check character
        sChkDigit = Mid(BC39_CharSet, iWeight + 1, 1)
        DataToEncode = DataToEncode & sChkDigit
        If cDisplayCheckDigit Then HRBarText = HRBarText & sChkDigit
     End If
     If cDisplayStartStop Then HRBarText = "*" & HRBarText & "*"
    'Add the start character
    EncodedData = BC39_Code(43) & "1"
    For iCntr = 1 To Len(DataToEncode)
        'Get each character one at a time
        sDataChar = (Mid(DataToEncode, iCntr, 1))
        EncodedData = EncodedData & BC39_Code(InStr(1, BC39_CharSet, sDataChar) - 1) & "1"
    Next iCntr
    'Add stop character
    EncodedData = EncodedData & BC39_Code(43) & "1"
    GetBarWidth EncodedData
End Function

Private Function Encode39Extd(DataToEncode)
    Dim sDataChar As String
    Dim sShiftChar As String
    Dim iCntr As Integer
    Dim iAscVal As Integer
    
    HRBarText = DataToEncode
    If cDisplayStartStop Then HRBarText = "*" & HRBarText & "*"
    EncodedData = BC39_Code(43) & "1"
    For iCntr = 1 To Len(DataToEncode)
        sDataChar = (Mid(DataToEncode, iCntr, 1))
        Select Case sDataChar
            Case "0" To "9", "A" To "Z", "-", ".", " " 'Unshifted characters
                sShiftChar = ""
                iAscVal = AscW(sDataChar)
            Case Chr(1) To Chr(26)      '$ shifted
                sShiftChar = "$"
                iAscVal = AscW(sDataChar) + 64
            Case "!" To ",", "/", ":"   '/ shifted
                sShiftChar = "/"
                iAscVal = AscW(sDataChar) + 32
            Case "a" To "z"             '+ shifted
                sShiftChar = "+"
                iAscVal = AscW(sDataChar) - 32
            Case Chr(0)                 '% shifted
                sShiftChar = "%"
                iAscVal = AscW(sDataChar) + 85
            Case Chr(27) To Chr(31)
                sShiftChar = "%"
                iAscVal = AscW(sDataChar) + 38
            Case ";" To "?"
                sShiftChar = "%"
                iAscVal = AscW(sDataChar) + 11
            Case "[" To "_"
                sShiftChar = "%"
                iAscVal = AscW(sDataChar) - 16
            Case "{" To "~", Chr(127)
                sShiftChar = "%"
                iAscVal = AscW(sDataChar) - 43
        End Select
        If sShiftChar <> "" Then EncodedData = EncodedData & BC39_Code(InStr(1, BC39_CharSet, sShiftChar) - 1) & "1"
        EncodedData = EncodedData & BC39_Code(InStr(1, BC39_CharSet, ChrW(iAscVal), vbTextCompare) - 1) & "1"
    Next iCntr
    'Add the stop character
    EncodedData = EncodedData & BC39_Code(43) & "1"
    GetBarWidth EncodedData
End Function

'=====================================
'CODE 128 FUNCTIONS
'=====================================

Private Function Encode128(DataToEncode As String)
    Dim iWeightTotal As Integer
    Dim iWeight As Integer
    Dim iCodeValue As Integer
    Dim iCheckDigit As Integer
    Dim iAscVal As Integer
    Dim iCntr As Integer
    Dim J As Integer
    Dim K As Integer
    Dim iDataLen As Integer
    Dim iCodeStart As Integer
    Dim sEncoding As String
    Dim sDataChar As String
    Dim rtn As Boolean
    
    iDataLen = Len(DataToEncode)
    'Set the human readable text
    HRBarText = FilterHRText(DataToEncode)
    
    'Select the character set A, B or C for the START character
    sDataChar = Left(DataToEncode, 1)
    iAscVal = AscW(sDataChar)
    If iAscVal < 32 Then iCodeStart = 103    'A
    If iAscVal > 31 And iAscVal < 127 Then iCodeStart = 104  'B
    If ((iDataLen > 3) And IsNumericOnly(Mid(DataToEncode, 1, 4))) Then iCodeStart = 105 'C
    Select Case iCodeStart
        Case 103
            sEncoding = "A"
            EncodedData = BC128_Code(103)
            iWeightTotal = 103
        Case 104
            sEncoding = "B"
            EncodedData = BC128_Code(104)
            iWeightTotal = 104
        Case 105
            sEncoding = "C"
            EncodedData = BC128_Code(105)
            iWeightTotal = 105
    End Select
    
    iWeight = 1
    For iCntr = 1 To iDataLen
        If ((iCntr < iDataLen - 2) And (IsNumericOnly(Mid(DataToEncode, iCntr, 4)))) _
            Or ((iCntr < iDataLen) And (sEncoding = "C") And (IsNumericOnly(Mid(DataToEncode, iCntr, 2)))) Then
            'Check to see if we have an odd number of numbers to encode.
            If sEncoding <> "C" Then
                J = iCntr
                K = 0
                Do While J <= iDataLen And IsNumericOnly(Mid(DataToEncode, J, 1))
                    J = J + 1
                    K = K + 1
                Loop
                If (K Mod 2) = 1 Then
                    'If so, stay in current set for 1 character to save space.
                    iAscVal = AscW(Mid(DataToEncode, iCntr, 1))
                    If sEncoding = "A" Then
                        If iAscVal > 31 Then iCodeValue = iAscVal - 32
                        If iAscVal < 32 Then iCodeValue = iAscVal + 64
                    ElseIf sEncoding = "B" Then
                        iCodeValue = iAscVal - 32
                    End If
                    EncodedData = EncodedData & BC128_Code(iCodeValue)
                    iWeightTotal = iWeightTotal + (iCodeValue * iWeight)
                    iCntr = iCntr + 1
                    iWeight = iWeight + 1
                End If
            End If
            'Switch to set C if not already in it
            If sEncoding <> "C" Then
                sEncoding = "C"
                EncodedData = EncodedData & BC128_Code(99)  'C switch
                iWeightTotal = iWeightTotal + (99 * iWeight)
                iWeight = iWeight + 1
            End If
            iAscVal = CInt((Mid(DataToEncode, iCntr, 2)))
            iWeightTotal = iWeightTotal + (iAscVal * iWeight)
            EncodedData = EncodedData & BC128_Code(iAscVal)
            iCntr = iCntr + 1
            iWeight = iWeight + 1
            
        'Check for switching to character set A
        ElseIf (iCntr <= iDataLen) And ((AscW(Mid(DataToEncode, iCntr, 1)) < 31) Or ((sEncoding = "A") _
                And (AscW(Mid(DataToEncode, iCntr, 1)) > 31 And (AscW(Mid(DataToEncode, iCntr, 1))) < 96))) Then
            'Switch to set A if not already in it
            If sEncoding <> "A" Then
                sEncoding = "A"
                EncodedData = EncodedData & BC128_Code(101) 'A switch
                iWeightTotal = iWeightTotal + (101 * iWeight)
                iWeight = iWeight + 1
            End If
            'Get the ASCII value of the next character
            iAscVal = AscW(Mid(DataToEncode, iCntr, 1))
            If iAscVal > 31 Then iCodeValue = iAscVal - 32
            If iAscVal < 32 Then iCodeValue = iAscVal + 64
            iWeightTotal = iWeightTotal + (iCodeValue * iWeight)
            EncodedData = EncodedData & BC128_Code(iCodeValue)
            iWeight = iWeight + 1
            
        'Check for switching to character set B
        ElseIf (iCntr <= iDataLen) And ((AscW(Mid(DataToEncode, iCntr, 1))) > 31 And _
                (AscW(Mid(DataToEncode, iCntr, 1)))) < 127 Then
            'Switch to set B if not already in it
            If sEncoding <> "B" Then
                EncodedData = EncodedData & BC128_Code(100) 'B switch
                sEncoding = "B"
                iWeightTotal = iWeightTotal + (100 * iWeight)
                iWeight = iWeight + 1
            End If
            'Get the ASCII value of the next character
            iAscVal = AscW(Mid(DataToEncode, iCntr, 1))
            iCodeValue = iAscVal - 32
            iWeightTotal = iWeightTotal + (iCodeValue * iWeight)
            EncodedData = EncodedData & BC128_Code(iCodeValue)
            iWeight = iWeight + 1
        End If
    Next iCntr
    'Calculate the check digit
    iCheckDigit = (iWeightTotal Mod 103)
    'Add the check digit, stop character and a wide termination bar
    EncodedData = EncodedData & BC128_Code(iCheckDigit) & BC128_Code(106) & "2"
    GetBarWidth EncodedData
End Function

Private Function Encode128a(DataToEncode As String)
    Dim iWeightTotal As Integer
    Dim iCodeValue As Integer
    Dim iCheckDigit As Integer
    Dim iAscVal As Integer
    Dim iCntr As Integer
    
    'Set the human readable text
    HRBarText = FilterHRText(DataToEncode)
    'Character set "A" start character
    EncodedData = BC128_Code(103)
    iWeightTotal = 103
    'Process each character in turn
    For iCntr = 1 To Len(DataToEncode)
        iAscVal = AscW(Mid(DataToEncode, iCntr, 1))
        If iAscVal > 31 Then iCodeValue = iAscVal - 32
        If iAscVal < 32 Then iCodeValue = iAscVal + 64
        iWeightTotal = iWeightTotal + (iCodeValue * iCntr)
        EncodedData = EncodedData & BC128_Code(iCodeValue)
     Next iCntr
     'Calculate the check digit
     iCheckDigit = (iWeightTotal Mod 103)
     'Add the check digit, stop character and a wide termination bar
     EncodedData = EncodedData & BC128_Code(iCheckDigit) & BC128_Code(106) & "2"
     GetBarWidth EncodedData
End Function

Private Function Encode128b(ByRef DataToEncode As String)
    Dim iWeightTotal As Integer
    Dim iCodeValue As Integer
    Dim iCheckDigit As Integer
    Dim iAscVal As Integer
    Dim iCntr As Integer
    
    'Set the human readable text
    HRBarText = DataToEncode
    'Character set "B" start character
    EncodedData = BC128_Code(104)
    iWeightTotal = 104
    'Process each character in turn
    For iCntr = 1 To Len(DataToEncode)
        iAscVal = AscW(Mid(DataToEncode, iCntr, 1))
        iCodeValue = iAscVal - 32
        iWeightTotal = iWeightTotal + (iCodeValue * iCntr)
        EncodedData = EncodedData & BC128_Code(iCodeValue)
     Next iCntr
     'Calculate the check digit
     iCheckDigit = (iWeightTotal Mod 103)
     'Add the check digit, stop character and the termination bar
     EncodedData = EncodedData & BC128_Code(iCheckDigit) & BC128_Code(106) & "2"
     GetBarWidth EncodedData
End Function

Private Function Encode128c(DataToEncode As String)
    Dim iWeightTotal As Integer
    Dim iWeight As Integer
    Dim iCheckDigit As Integer
    Dim iAscVal As Integer
    Dim iCntr As Integer
    
    'Code "C" must have an even number of numeric characters.
    'Pad out with a leading "0" if not.
    If (Len(DataToEncode) Mod 2) = 1 Then DataToEncode = "0" & DataToEncode
    'Set the human readable text
    HRBarText = DataToEncode
    'Character set "C" start character
    iWeightTotal = 105
    EncodedData = BC128_Code(105)
    iWeight = 1
    'Process each character in turn
    For iCntr = 1 To Len(DataToEncode) Step 2
        iAscVal = CInt(Mid(DataToEncode, iCntr, 2))
        iWeightTotal = iWeightTotal + (iAscVal * iWeight)
        EncodedData = EncodedData & BC128_Code(iAscVal)
        iWeight = iWeight + 1
    Next iCntr
    'Calculate the check digit
    iCheckDigit = iWeightTotal Mod 103
    'Add the check digit, stop character and the termination bar
    EncodedData = EncodedData & BC128_Code(iCheckDigit) & BC128_Code(106) & "2"
    GetBarWidth EncodedData
End Function

'=====================================
'CODE 93 FUNCTIONS
'=====================================

Private Function Encode93(DataToEncode As String)
    Dim iCntr As Integer
    Dim iCodeValue As Integer
    Dim iCWCntr As Integer
    Dim iKWCntr As Integer
    Dim CWeight As Integer
    Dim KWeight As Integer
    Dim CCheckSum As Integer
    Dim KCheckSum As Integer
    Dim sDataChar As String * 1
    
    HRBarText = DataToEncode
    If cDisplayStartStop Then HRBarText = "*" & HRBarText & "*"
    'Calculate "C" and "K" check digits
    iCWCntr = 1
    iKWCntr = 2
    For iCntr = Len(DataToEncode) To 1 Step -1
        sDataChar = (Mid(DataToEncode, iCntr, 1))
        iCodeValue = InStr(1, BC93_CharSet, sDataChar) - 1
        CWeight = CWeight + (iCodeValue * iCWCntr)
        iCWCntr = iCWCntr + 1
        If iCWCntr = 21 Then iCWCntr = 1
        KWeight = KWeight + (iCodeValue * iKWCntr)
        iKWCntr = iKWCntr + 1
        If iKWCntr = 16 Then iKWCntr = 1
    Next iCntr
    CCheckSum = CWeight Mod 47
    KWeight = KWeight + CCheckSum
    KCheckSum = KWeight Mod 47
    'Add start character
    EncodedData = BC93_Code(47)
    For iCntr = 1 To Len(DataToEncode)
        'Get each character one at a time
        sDataChar = (Mid(DataToEncode, iCntr, 1))
        EncodedData = EncodedData & BC93_Code(InStr(1, BC93_CharSet, sDataChar) - 1)
    Next iCntr
    'Add the "C" and "K" checksums,stop character and termination bar
    EncodedData = EncodedData & BC93_Code(CCheckSum) & BC93_Code(KCheckSum) & BC93_Code(47) & "1"
    GetBarWidth EncodedData
End Function

'=====================================
'CODABAR FUNCTION
'=====================================

Private Function EncodeCodaBar(DataToEncode As String)
    Dim iCntr As Integer
    Dim iBarCnt As Integer
    Dim sDataChar As String
    
    HRBarText = DataToEncode
    
    'Add "A" start character
    EncodedData = Codabar_Code(16) & "1"
    For iCntr = 1 To Len(DataToEncode)
        'Get each character one at a time
        sDataChar = (Mid(DataToEncode, iCntr, 1))
        EncodedData = EncodedData & Codabar_Code(InStr(1, Codabar_CharSet, sDataChar) - 1) & "1"
    Next iCntr
    'Add the "B" stop character
    EncodedData = EncodedData & Codabar_Code(17) & "1"
    GetBarWidth EncodedData
End Function

'=====================================
'MSI FUNCTIONS
'=====================================

Private Function EncodeMSI(DataToEncode As String)
    Dim iCntr As Integer
    Dim sDataChar As String * 1
    Dim iBarCnt As Integer
    Dim sOddNumbers As String
    Dim OddDigit As Boolean
    Dim EvenNumberTotal As Integer
    Dim OddNumberTotal As Integer
    Dim sOddNumberProduct As String
    Dim sChkDigit As String * 1
    Dim iProduct As Integer
    Dim iCarry As Integer

    HRBarText = DataToEncode
    'MSI supports 3 checksum methods. MOD10,MOD10 x 2 and MOD11
    'We will use the MOD10 checksum
    '<<<< Calculate Check Digit >>>>
    'To create the check digit follow these steps:
    '1)Starting from the units position, create a new number with all of the odd
    '  position digits in their original sequence.
    '2)Multiply this new number by 2.
    '3)Add all of the digits of the product from step two.
    '4)Add all of the digits not used in step one to the result in step three.
    '5)Determine the smallest number which when added to the result in step four
    '  will result in a multiple of 10. This is the check character.
    
    'Step 1 -- Create a new number of the odd position digits starting from the right and going left, but store the
    'digits from left to right.
    'Create the odd position number & prepare for Step 4 by getting the sum of all even position characters
    sOddNumbers = ""
    OddDigit = True
    EvenNumberTotal = 0
    For iCntr = Len(DataToEncode) To 1 Step -1
        If OddDigit = True Then
            sOddNumbers = Mid(DataToEncode, iCntr, 1) & sOddNumbers
            OddDigit = False
        Else
            EvenNumberTotal = EvenNumberTotal + CInt(Mid(DataToEncode, iCntr, 1))
            OddDigit = True
        End If
    Next iCntr
    'Step 2 -- Multiply this new number by 2.
    'Convert string to long, * 2 and convert back to string will cause overflow errors on long barcodes.
    'Use this method to perform the muliplication to avoid the chance of an overflow error.
    iCarry = 0
    For iCntr = Len(sOddNumbers) To 1 Step -1
        iProduct = (CInt(Mid(sOddNumbers, iCntr, 1)) * 2) + iCarry
        If iProduct > 9 Then
            iCarry = 1
            sOddNumberProduct = CStr(iProduct Mod 10) & sOddNumberProduct
        Else
            iCarry = 0
            sOddNumberProduct = CStr(iProduct) & sOddNumberProduct
        End If
    Next iCntr
    If iCarry = 1 Then sOddNumberProduct = CStr(iCarry) & sOddNumberProduct
    
    'Step 3 -- Add all of the digits of the product from step two.
    
    OddNumberTotal = 0
    For iCntr = 1 To Len(sOddNumberProduct)
        OddNumberTotal = OddNumberTotal + CInt(Mid(sOddNumberProduct, iCntr, 1))
    Next iCntr
    'Step 4 -- Add all of the digits not used in step one to the result in step three.
    'We will store the result in OddNumberSum just so we don't have to create another variable
    OddNumberTotal = OddNumberTotal + EvenNumberTotal
    'Step 5 -- Determine the smallest number which when added to the result in step four
    'will result in a multiple of 10. This is the check character.
    OddNumberTotal = OddNumberTotal Mod 10
    If OddNumberTotal <> 0 Then
        sChkDigit = CStr(10 - OddNumberTotal)
    Else
        sChkDigit = "0"
    End If
    If cDisplayCheckDigit Then HRBarText = HRBarText & sChkDigit
    DataToEncode = DataToEncode & sChkDigit
    If cDisplayCheckDigit = True Then HRBarText = DataToEncode
    'Add the start character
    EncodedData = MSI_Code(10)
    For iCntr = 1 To Len(DataToEncode)
        'Get each character one at a time
        sDataChar = (Mid(DataToEncode, iCntr, 1))
        EncodedData = EncodedData & MSI_Code(InStr(1, MSI_Charset, sDataChar) - 1)
    Next iCntr
    'Add the stop character
    EncodedData = EncodedData & MSI_Code(11)
    GetBarWidth EncodedData
End Function

'=====================================
'CODE 11 FUNCTION
'=====================================
Private Function Encode11(DataToEncode As String)
    Dim iCntr As Integer
    Dim sDataChar As String * 1
    Dim iCodeValue As Integer
    Dim iCWCntr As Integer
    Dim iKWCntr As Integer
    Dim CWeight As Integer
    Dim KWeight As Integer
    Dim CCheckSum As Integer
    Dim KCheckSum As Integer
    Dim iBarCnt As Integer
    
    HRBarText = DataToEncode
    If cUseCheckDigit Then
        'Calculate the "C" & "K" checksums
        iCWCntr = 1
        iKWCntr = 2
        For iCntr = Len(DataToEncode) To 1 Step -1
            sDataChar = (Mid(DataToEncode, iCntr, 1))
            iCodeValue = InStr(1, BC11_Charset, sDataChar) - 1
            CWeight = CWeight + (iCodeValue * iCWCntr)
            iCWCntr = iCWCntr + 1
            If iCWCntr = 11 Then iCWCntr = 1
            KWeight = KWeight + (iCodeValue * iKWCntr)
            iKWCntr = iKWCntr + 1
            If iKWCntr = 10 Then iKWCntr = 1
        Next iCntr
        CCheckSum = CWeight Mod 11
        KWeight = KWeight + CCheckSum
        KCheckSum = KWeight Mod 11
        'Use the "C" check digit for data up to 10 characters and the "C" & "K" for longer
        DataToEncode = DataToEncode & Mid(BC11_Charset, CCheckSum + 1, 1)
        If Len(DataToEncode) > 11 Then DataToEncode = DataToEncode & Mid(BC11_Charset, KCheckSum + 1, 1)
        If cDisplayCheckDigit = True Then HRBarText = DataToEncode
    End If
    'Add start character and the quite space
    EncodedData = BC11_Code(11) & "1"
    For iCntr = 1 To Len(DataToEncode)
        'Get each character one at a time
        sDataChar = (Mid(DataToEncode, iCntr, 1))
        EncodedData = EncodedData & BC11_Code(InStr(1, BC11_Charset, sDataChar) - 1) & "1"
    Next iCntr
    'Add stop character
    EncodedData = EncodedData & BC11_Code(11) & "1"
    GetBarWidth EncodedData
End Function

'=====================================
'CODE 2of5 FUNCTIONS
'=====================================
Private Function Encode2of5(DataToEncode As String)
    Dim iCntr As Integer
    Dim sDataChar As String * 1
    Dim iBarCnt As Integer
    
    HRBarText = DataToEncode
    'Add start character and the quite space
    EncodedData = BC2of5_Code(10)
    For iCntr = 1 To Len(DataToEncode)
        'Get each character one at a time
        sDataChar = (Mid(DataToEncode, iCntr, 1))
        EncodedData = EncodedData & BC2of5_Code(InStr(1, BC2of5_Charset, sDataChar) - 1)
    Next iCntr
    'Add stop character
    EncodedData = EncodedData & BC2of5_Code(11)
    GetBarWidth EncodedData
End Function

Private Function EncodeI2of5(DataToEncode As String)
    Dim iCntr As Integer
    Dim iWeight As Integer
    Dim sChkDigit As String * 1
    Dim TotalWeight As Integer
    Dim sDataChar As String * 1
    Dim sCodeData1 As String
    Dim sCodeData2 As String
    Dim iBarCnt As Integer
    
    HRBarText = DataToEncode
    If cUseCheckDigit Then
        'Calculate the check digit
        iWeight = 3
        For iCntr = Len(DataToEncode) To 1 Step -1
            TotalWeight = TotalWeight + (CInt(Mid(DataToEncode, iCntr, 1)) * iWeight)
            If iWeight = 3 Then iWeight = 1 Else iWeight = 3
        Next iCntr
        TotalWeight = TotalWeight Mod 10
        If TotalWeight <> 0 Then
            sChkDigit = CStr(10 - TotalWeight)
        Else
            sChkDigit = "0"
        End If
        DataToEncode = DataToEncode & sChkDigit
        If cDisplayCheckDigit Then HRBarText = HRBarText & sChkDigit
    End If
    'Code "I2of5" must have an even number of numeric characters.
    'Pad out with a leading "0" if not.
    If (Len(DataToEncode) Mod 2) = 1 Then
        DataToEncode = "0" & DataToEncode
        HRBarText = "0" & HRBarText
    End If
    
    'Add start character and the quite space
    EncodedData = BCI2of5_code(10)
    For iCntr = 1 To Len(DataToEncode) Step 2
        'Get each character pair at a time
        sDataChar = (Mid(DataToEncode, iCntr, 1))
        sCodeData1 = BCI2of5_code(InStr(1, BC2of5_Charset, sDataChar) - 1)
        sDataChar = (Mid(DataToEncode, iCntr + 1, 1))
        sCodeData2 = BCI2of5_code(InStr(1, BC2of5_Charset, sDataChar) - 1)
        'Interleave the encoded data
        For iBarCnt = 1 To Len(sCodeData1)
            EncodedData = EncodedData & Mid(sCodeData1, iBarCnt, 1) & Mid(sCodeData2, iBarCnt, 1)
        Next iBarCnt
    Next iCntr
    'Add stop character
    EncodedData = EncodedData & BCI2of5_code(11)
    GetBarWidth EncodedData
End Function

Private Sub GetBarWidth(sData As String)
    'Calculate the width of the barcode
    Dim iCntr As Integer
    Dim iBarCnt As Integer
    Dim iBW As Integer
    
    NarrowBar = cBarX * cBarMultiplier
    For iCntr = 1 To Len(EncodedData)
        iBW = CInt(Mid(EncodedData, iCntr, 1))
        If iBW > 1 Then iBW = iBW * cBarRatio
        iBarCnt = iBarCnt + iBW
    Next iCntr
    BarCodeWidth = NarrowBar * iBarCnt
End Sub

Private Function FilterHRText(fText As String) As String
    'Filter out any non printable characters of the
    '128/128a character set for the human readable text.
    Dim iCntr As Integer
    Dim sDataChar As String
    Dim oText As String
    
    For iCntr = 1 To Len(fText)
        sDataChar = Mid(fText, iCntr, 1)
        If Asc(sDataChar) > 31 Then oText = oText & sDataChar
    Next iCntr
    FilterHRText = oText
End Function

Private Function IsNumericOnly(sNumericText As String) As Boolean
    'Test if a string contains only numeric characters 0-9
     Dim iCntr As Integer
     Dim iAscVal As Integer

     For iCntr = 1 To Len(sNumericText)
        iAscVal = AscW(Mid(sNumericText, iCntr, 1))
        If iAscVal < 48 Or iAscVal > 57 Then
             IsNumericOnly = False
             Exit Function
        End If
     Next iCntr
     IsNumericOnly = True
End Function

Private Sub Class_Initialize()
    'Initialize some default values
    'Barcode
    cSymbology = 0
    cBarX = 1
    cBarMultiplier = 1
    cBarRatio = 1
    cUseCheckDigit = False
    cBarRotation = 0
    cBarCodeTop = 1
    cBarCodeLeft = 1
    cBarCodeHeight = 45
    cBarTextGap = 0
    cBarCodeData = ""
    
    'Barcode text
    cHumanReadable = True
    cBarTextPosition = TP_UNDER
    cBarTextAlign = ALIGN_CENTRE
    cDisplayStartStop = False
    cDisplayCheckDigit = False
    cBarTextFont.FontName = "Tahoma"
    cBarTextFont.FontSize = 8
    cBarTextFont.FontBold = False
    cBarTextFont.FontItalic = False
    
    'Errors
    cErrorNumber = 0
    cErrDescription = ""
    
    'Barcode characters
    Initialise_BC
End Sub

Private Sub Initialise_BC()
    '=====================================
    'CODE 39 CHARACTER SET
    '=====================================
    'Valid characters for Code 39
    BC39_CharSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
    
    'Bar pattens for each valid character
    BC39_Code(0) = "111221211": BC39_Code(1) = "211211112"
    BC39_Code(2) = "112211112": BC39_Code(3) = "212211111"
    BC39_Code(4) = "111221112": BC39_Code(5) = "211221111"
    BC39_Code(6) = "112221111": BC39_Code(7) = "111211212"
    BC39_Code(8) = "211211211": BC39_Code(9) = "112211211"
    BC39_Code(10) = "211112112": BC39_Code(11) = "112112112"
    BC39_Code(12) = "212112111": BC39_Code(13) = "111122112"
    BC39_Code(14) = "211122111": BC39_Code(15) = "112122111"
    BC39_Code(16) = "111112212": BC39_Code(17) = "211112211"
    BC39_Code(18) = "112112211": BC39_Code(19) = "111122211"
    BC39_Code(20) = "211111122": BC39_Code(21) = "112111122"
    BC39_Code(22) = "212111121": BC39_Code(23) = "111121122"
    BC39_Code(24) = "211121121": BC39_Code(25) = "112121121"
    BC39_Code(26) = "111111222": BC39_Code(27) = "211111221"
    BC39_Code(28) = "112111221": BC39_Code(29) = "111121221"
    BC39_Code(30) = "221111112": BC39_Code(31) = "122111112"
    BC39_Code(32) = "222111111": BC39_Code(33) = "121121112"
    BC39_Code(34) = "221121111": BC39_Code(35) = "122121111"
    BC39_Code(36) = "121111212": BC39_Code(37) = "221111211"
    BC39_Code(38) = "122111211": BC39_Code(39) = "121212111"
    BC39_Code(40) = "121211121": BC39_Code(41) = "121112121"
    BC39_Code(42) = "111212121": BC39_Code(43) = "121121211"
    
    '=====================================
    'CODE 93 CHARACTER SET
    '=====================================
    'Valid characters for Code 93
    BC93_CharSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
    
    BC93_Code(0) = "131112": BC93_Code(1) = "111213"
    BC93_Code(2) = "111312": BC93_Code(3) = "111411"
    BC93_Code(4) = "121113": BC93_Code(5) = "121212"
    BC93_Code(6) = "121311": BC93_Code(7) = "111114"
    BC93_Code(8) = "131211": BC93_Code(9) = "141111"
    BC93_Code(10) = "211113": BC93_Code(11) = "211212"
    BC93_Code(12) = "211311": BC93_Code(13) = "221112"
    BC93_Code(14) = "221211": BC93_Code(15) = "231111"
    BC93_Code(16) = "112113": BC93_Code(17) = "112212"
    BC93_Code(18) = "112311": BC93_Code(19) = "122112"
    BC93_Code(20) = "132111": BC93_Code(21) = "111123"
    BC93_Code(22) = "111222": BC93_Code(23) = "111321"
    BC93_Code(24) = "121122": BC93_Code(25) = "131121"
    BC93_Code(26) = "212112": BC93_Code(27) = "212211"
    BC93_Code(28) = "211122": BC93_Code(29) = "211221"
    BC93_Code(30) = "221121": BC93_Code(31) = "222111"
    BC93_Code(32) = "112122": BC93_Code(33) = "112221"
    BC93_Code(34) = "122121": BC93_Code(35) = "123111"
    BC93_Code(36) = "121131": BC93_Code(37) = "311112"
    BC93_Code(38) = "311211": BC93_Code(39) = "321111"
    BC93_Code(40) = "112131": BC93_Code(41) = "113121"
    BC93_Code(42) = "211131": BC93_Code(43) = "121221"
    BC93_Code(44) = "312111": BC93_Code(45) = "311121"
    BC93_Code(46) = "122211": BC93_Code(47) = "111141"
    
    '=====================================
    'CODE 128 CHARACTER SETS
    '=====================================
    
    'Valid characters for Code 128
    BC128A_CharSet = " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_" _
                    & Chr(0) & Chr(1) & Chr(2) & Chr(3) & Chr(4) & Chr(5) & Chr(6) & Chr(7) & Chr(8) _
                    & Chr(9) & Chr(10) & Chr(11) & Chr(12) & Chr(13) & Chr(14) & Chr(15) & Chr(16) _
                    & Chr(17) & Chr(18) & Chr(19) & Chr(20) & Chr(21) & Chr(22) & Chr(23) & Chr(24) _
                    & Chr(25) & Chr(26) & Chr(27) & Chr(28) & Chr(29) & Chr(30) & Chr(31)
                                        
    BC128B_CharSet = " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_" _
                    & "`abcdefghijklmnopqrstuvwxyz{|}~" & Chr(127)
    
    BC128C_CharSet = "0123456789"
    
    'Bar pattens for each valid character
    BC128_Code(0) = "212222": BC128_Code(1) = "222122": BC128_Code(2) = "222221"
    BC128_Code(3) = "121223": BC128_Code(4) = "121322": BC128_Code(5) = "131222"
    BC128_Code(6) = "122213": BC128_Code(7) = "122312": BC128_Code(8) = "132212"
    BC128_Code(9) = "221213": BC128_Code(10) = "221312": BC128_Code(11) = "231212"
    BC128_Code(12) = "112232": BC128_Code(13) = "122132": BC128_Code(14) = "122231"
    BC128_Code(15) = "113222": BC128_Code(16) = "123122": BC128_Code(17) = "123221"
    BC128_Code(18) = "223211": BC128_Code(19) = "221132": BC128_Code(20) = "221231"
    BC128_Code(21) = "213212": BC128_Code(22) = "223112": BC128_Code(23) = "312131"
    BC128_Code(24) = "311222": BC128_Code(25) = "321122": BC128_Code(26) = "321221"
    BC128_Code(27) = "312212": BC128_Code(28) = "322112": BC128_Code(29) = "322211"
    BC128_Code(30) = "212123": BC128_Code(31) = "212321": BC128_Code(32) = "232121"
    BC128_Code(33) = "111323": BC128_Code(34) = "131123": BC128_Code(35) = "131321"
    BC128_Code(36) = "112313": BC128_Code(37) = "142113": BC128_Code(38) = "132311"
    BC128_Code(39) = "211313": BC128_Code(40) = "231113": BC128_Code(41) = "231311"
    BC128_Code(42) = "112133": BC128_Code(43) = "112331": BC128_Code(44) = "132131"
    BC128_Code(45) = "113123": BC128_Code(46) = "113321": BC128_Code(47) = "133121"
    BC128_Code(48) = "313121": BC128_Code(49) = "211331": BC128_Code(50) = "231131"
    BC128_Code(51) = "213113": BC128_Code(52) = "213311": BC128_Code(53) = "213131"
    BC128_Code(54) = "311123": BC128_Code(55) = "311321": BC128_Code(56) = "331121"
    BC128_Code(57) = "312113": BC128_Code(58) = "312311": BC128_Code(59) = "332111"
    BC128_Code(60) = "324111": BC128_Code(61) = "221411": BC128_Code(62) = "431111"
    BC128_Code(63) = "111224": BC128_Code(64) = "111422": BC128_Code(65) = "121124"
    BC128_Code(66) = "121421": BC128_Code(67) = "141122": BC128_Code(68) = "141221"
    BC128_Code(69) = "112214": BC128_Code(70) = "112412": BC128_Code(71) = "122114"
    BC128_Code(72) = "122411": BC128_Code(73) = "142112": BC128_Code(74) = "142211"
    BC128_Code(75) = "241211": BC128_Code(76) = "221114": BC128_Code(77) = "413111"
    BC128_Code(78) = "241112": BC128_Code(79) = "134111": BC128_Code(80) = "111242"
    BC128_Code(81) = "121142": BC128_Code(82) = "121241": BC128_Code(83) = "114212"
    BC128_Code(84) = "124112": BC128_Code(85) = "124211": BC128_Code(86) = "411212"
    BC128_Code(87) = "421112": BC128_Code(88) = "421211": BC128_Code(89) = "212141"
    BC128_Code(90) = "214121": BC128_Code(91) = "412121": BC128_Code(92) = "111143"
    BC128_Code(93) = "111341": BC128_Code(94) = "131141": BC128_Code(95) = "114113"
    BC128_Code(96) = "114311": BC128_Code(97) = "411112": BC128_Code(98) = "411311"
    BC128_Code(99) = "113141": BC128_Code(100) = "114131": BC128_Code(101) = "311141"
    BC128_Code(102) = "411131": BC128_Code(103) = "211412": BC128_Code(104) = "211214"
    BC128_Code(105) = "211232": BC128_Code(106) = "233111"
    
    '=====================================
    'CODABAR CHARACTER SET
    '=====================================
    
    Codabar_CharSet = "0123456789-$:/.+"
    
    Codabar_Code(0) = "1111122": Codabar_Code(1) = "1111221"
    Codabar_Code(2) = "1112112": Codabar_Code(3) = "2211111"
    Codabar_Code(4) = "1121121": Codabar_Code(5) = "2111121"
    Codabar_Code(6) = "1211112": Codabar_Code(7) = "1211211"
    Codabar_Code(8) = "1221111": Codabar_Code(9) = "2112111"
    Codabar_Code(10) = "1112211": Codabar_Code(11) = "1122111"
    Codabar_Code(12) = "2111212": Codabar_Code(13) = "2121112"
    Codabar_Code(14) = "2121211": Codabar_Code(15) = "1121212"
    Codabar_Code(16) = "1122121": Codabar_Code(17) = "1212112"
    Codabar_Code(18) = "1112122": Codabar_Code(19) = "1112221"
    
    '=====================================
    'MSI CHARACTER SET
    '=====================================
    
    MSI_Charset = "0123456789"
    
    MSI_Code(0) = "12121212": MSI_Code(1) = "12121221"
    MSI_Code(2) = "12122112": MSI_Code(3) = "12122121"
    MSI_Code(4) = "12211212": MSI_Code(5) = "12211221"
    MSI_Code(6) = "12212112": MSI_Code(7) = "12212121"
    MSI_Code(8) = "21121212": MSI_Code(9) = "21121221"
    MSI_Code(10) = "21": MSI_Code(11) = "121"
    
    
    '=====================================
    'CODE 11 CHARACTER SET
    '=====================================
    
    BC11_Charset = "0123456789-"
    
    BC11_Code(0) = "11113": BC11_Code(1) = "31113"
    BC11_Code(2) = "13113": BC11_Code(3) = "33111"
    BC11_Code(4) = "11313": BC11_Code(5) = "31311"
    BC11_Code(6) = "13311": BC11_Code(7) = "11133"
    BC11_Code(8) = "31131": BC11_Code(9) = "31111"
    BC11_Code(10) = "11311": BC11_Code(11) = "11331"
    
    '=====================================
    'CODE 2of5 CHARACTER SET
    '=====================================
    
    BC2of5_Charset = "0123456789"
    
    'Standard bar pattens
    BC2of5_Code(0) = "1111313111": BC2of5_Code(1) = "3111111131"
    BC2of5_Code(2) = "1131111131": BC2of5_Code(3) = "3131111111"
    BC2of5_Code(4) = "1111311131": BC2of5_Code(5) = "3111311111"
    BC2of5_Code(6) = "1131311111": BC2of5_Code(7) = "1111113131"
    BC2of5_Code(8) = "3111113111": BC2of5_Code(9) = "1131113111"
    BC2of5_Code(10) = "313111": BC2of5_Code(11) = "31113"
    
    'Interleave bar pattens
    BCI2of5_code(0) = "11221": BCI2of5_code(1) = "21112"
    BCI2of5_code(2) = "12112": BCI2of5_code(3) = "22111"
    BCI2of5_code(4) = "11212": BCI2of5_code(5) = "21211"
    BCI2of5_code(6) = "12211": BCI2of5_code(7) = "11122"
    BCI2of5_code(8) = "21121": BCI2of5_code(9) = "12121"
    BCI2of5_code(10) = "1111": BCI2of5_code(11) = "211"
    
End Sub




