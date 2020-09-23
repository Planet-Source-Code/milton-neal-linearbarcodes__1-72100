VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   8145
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8910
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboXFactor 
      Height          =   315
      Left            =   7920
      TabIndex        =   40
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox chkCD 
      Caption         =   "Use Check Digit"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   38
      ToolTipText     =   "Use a Check Digit if optional for code."
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Barcode Color ...."
      Height          =   315
      Left            =   7200
      TabIndex        =   35
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Human Readable"
      Height          =   1815
      Left            =   240
      TabIndex        =   32
      Top             =   5280
      Width           =   2175
      Begin VB.CheckBox chkDCD 
         Caption         =   "Display Check Digit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkDSS 
         Caption         =   "Display Stop/Start"
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton optHR 
         Caption         =   "Barcode and Text"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optHR 
         Caption         =   "Barcode Only"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rotation"
      Height          =   1815
      Left            =   4320
      TabIndex        =   27
      Top             =   5280
      Width           =   2055
      Begin VB.OptionButton optRotation 
         Caption         =   "Upside Down"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton optRotation 
         Caption         =   "Sideways Down"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton optRotation 
         Caption         =   "Sideways UP"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optRotation 
         Caption         =   "Normal"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Font ...."
      Height          =   315
      Left            =   6360
      TabIndex        =   26
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtGap 
      Height          =   315
      Left            =   7920
      TabIndex        =   25
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Text Alignment"
      Height          =   1815
      Left            =   2520
      TabIndex        =   19
      Top             =   5280
      Width           =   1695
      Begin VB.OptionButton optTxtAlign 
         Caption         =   "Full Width"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton optTxtAlign 
         Caption         =   "Right"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optTxtAlign 
         Caption         =   "Left"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optTxtAlign 
         Caption         =   "Centred"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text Position"
      Height          =   975
      Left            =   6480
      TabIndex        =   16
      Top             =   5280
      Width           =   1935
      Begin VB.OptionButton optTxtPos 
         Caption         =   "Under Barcode"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optTxtPos 
         Caption         =   "Above Barcode"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.ComboBox cboBarRatio 
      Height          =   315
      Left            =   7920
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox cboMultiplier 
      Height          =   315
      Left            =   7920
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   7440
      Width           =   900
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "&Draw"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   7440
      Width           =   900
   End
   Begin VB.TextBox txtHeight 
      Height          =   315
      Left            =   7920
      TabIndex        =   4
      Text            =   "12"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtY 
      Height          =   315
      Left            =   7920
      TabIndex        =   3
      Text            =   "5"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtX 
      Height          =   315
      Left            =   7920
      TabIndex        =   2
      Text            =   "5"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtBarcode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   1
      Top             =   3600
      Width           =   2415
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   39
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picWorkspace 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label lblImage 
      AutoSize        =   -1  'True
      Caption         =   "Bitmap Image"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   7080
      TabIndex        =   42
      Top             =   6480
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar X Width"
      Height          =   195
      Left            =   6360
      TabIndex        =   41
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar/Text Gap [mm]"
      Height          =   195
      Left            =   6360
      TabIndex        =   24
      Top             =   1440
      Width           =   1350
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Ratio:"
      Height          =   195
      Left            =   6360
      TabIndex        =   15
      Top             =   2520
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multiplier"
      Height          =   195
      Left            =   6360
      TabIndex        =   14
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Height: [mm]"
      Height          =   195
      Left            =   6360
      TabIndex        =   13
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y: [mm]"
      Height          =   195
      Left            =   6360
      TabIndex        =   12
      Top             =   720
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X: [mm]"
      Height          =   195
      Left            =   6360
      TabIndex        =   11
      Top             =   360
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode Data:"
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   75
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveBMP 
         Caption         =   "Save as Bitmap"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuBarCode 
      Caption         =   "Barcode"
      Begin VB.Menu mnucode128Auto 
         Caption         =   "Code 128 Auto"
      End
      Begin VB.Menu mnucode128A 
         Caption         =   "Code 128A"
      End
      Begin VB.Menu mnuCode128b 
         Caption         =   "Code 128B"
      End
      Begin VB.Menu mnuCode128c 
         Caption         =   "Code 128C"
      End
      Begin VB.Menu mnuCodeSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCode39 
         Caption         =   "Code 39"
      End
      Begin VB.Menu mnuCode39Extd 
         Caption         =   "Code 39 Full ASCII"
      End
      Begin VB.Menu mnuCodeSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCode2of5 
         Caption         =   "Code 2of5"
      End
      Begin VB.Menu mnuCodeI2of5 
         Caption         =   "Code I2of5"
      End
      Begin VB.Menu mnuCodeSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCode11 
         Caption         =   "Code 11"
      End
      Begin VB.Menu mnuCode93 
         Caption         =   "Code 93"
      End
      Begin VB.Menu mnuCodeMSI 
         Caption         =   "MSI (Plessy)"
      End
      Begin VB.Menu mnuCodeCodabar 
         Caption         =   "Codabar"
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Dim BC As clsLinearBarCodes
Dim Ratio As Single
Dim fname As String
Dim fsize As Single
Dim fbold As Boolean
Dim fitalic As Boolean

Private Sub cboBarRatio_Click()
    Ratio = CSng(Left(Me.cboBarRatio.Text, 1) / Right(Me.cboBarRatio, 1))
    BC.BarRatio = Ratio
End Sub

Private Sub cmdFont_Click()
    CommonDialog1.FontName = Me.txtBarcode.FontName
    ' Set Cancel to True
    CommonDialog1.CancelError = True
    On Error GoTo errHandler
    ' Set the Flags property
    CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
    ' Display the Font dialog box
    CommonDialog1.ShowFont
    fname = CommonDialog1.FontName
    fsize = CommonDialog1.FontSize
    fbold = CommonDialog1.FontBold
    fitalic = CommonDialog1.FontItalic
    Me.txtBarcode.Font.Name = CommonDialog1.FontName
    Me.txtBarcode.Font.Bold = CommonDialog1.FontBold
    Me.txtBarcode.Font.Italic = CommonDialog1.FontItalic
    cmdDraw_Click
Exit Sub
errHandler:
  ' User pressed the Cancel button
  Exit Sub

End Sub

Private Sub cmdColor_Click()
On Error GoTo errHandler
With CommonDialog1
    'Set Cancel to True
    .CancelError = True
    'Set the Flags property
    .Flags = cdlCCRGBInit 'Or cdlCCFullOpen
    'Display the Color Dialog box
    .ShowColor
    BC.BarColor = .Color
End With
cmdDraw_Click
errHandler:
' User pressed the Cancel button
End Sub

Private Sub cmdDraw_Click()
    Dim MilsToPixels As Single
    Dim th As Long
    Dim tw As Long
    Dim pr As Single
    Dim br As Single

    On Error GoTo Err_Handler

    MilsToPixels = 1440 / Screen.TwipsPerPixelX
    pr = CSng(Me.cboXFactor.Text) * MilsToPixels
    picWorkspace.Cls
    With BC
        .BarcodeOutput = Me.picWorkspace
        .BarCodeX = CLng(Me.txtX.Text) * (56.7 / Screen.TwipsPerPixelX)
        .BarCodeY = CLng(Me.txtY.Text) * (56.7 / Screen.TwipsPerPixelY)
        .BarXFactor = pr
        .BarRatio = Ratio
        .BarMultiplier = CInt(Me.cboMultiplier.Text)
        .BarcodeHeight = CLng(Me.txtHeight.Text) * (56.7 / Screen.TwipsPerPixelY)
        .BarTextGap = CInt(Me.txtGap.Text) * (56.7 / Screen.TwipsPerPixelY)
        .BarTextFont fname, fsize, fbold, fitalic
        .BarCodeData = txtBarcode.Text
        .DrawBarCode
        th = .TotalBarHeight
        tw = .TotalBarWidth
    Call Rectangle(Me.picWorkspace.hdc, .BarCodeX - 2, .BarCodeY - 2, .BarCodeX + tw + 2, .BarCodeY + th + 2)
    End With
    Me.picWorkspace.Refresh
    Exit Sub
    
Err_Handler:
MsgBox "Error Number: " & CStr(Err.Number) & vbCrLf _
                        & "Description: " & Err.Description & vbCrLf _
                        & "Error Source: " & Err.Source, vbOKOnly + vbCritical, "Error"
End Sub

Private Sub cmdPrint_Click()
    Dim MilsToPixels As Single
    Dim pr As Single
    Dim br As Single
    
    On Error GoTo Err_Handler

    MilsToPixels = 1440 / Printer.TwipsPerPixelX
    pr = CSng(Me.cboXFactor.Text) * MilsToPixels

    With BC
        .BarcodeOutput = Printer
        .BarCodeX = CLng(Me.txtX.Text) * (56.7 / Printer.TwipsPerPixelX)
        .BarCodeY = CLng(Me.txtY.Text) * (56.7 / Printer.TwipsPerPixelY)
        .BarXFactor = Int(pr)
        .BarRatio = Ratio
        .BarMultiplier = CInt(Me.cboMultiplier.Text)
        ' Set the height needed
        .BarcodeHeight = CLng(Me.txtHeight.Text) * 56.7 / Printer.TwipsPerPixelY
        .BarTextGap = CInt(Me.txtGap.Text) * (56.7 / Printer.TwipsPerPixelY)
        .BarTextFont fname, fsize, fbold, fitalic
        .BarCodeData = txtBarcode.Text
        Printer.Print ""
        .DrawBarCode
    End With
    Printer.EndDoc
    Exit Sub
    
Err_Handler:
MsgBox "Error Number: " & CStr(Err.Number) & vbCrLf _
                        & "Description: " & Err.Description & vbCrLf _
                        & "Error Source: " & Err.Source, vbOKOnly + vbCritical, "Error"
End Sub

Private Sub mnuFileSaveBMP_Click()
    'Dim MilsToPixels As Single
    Dim th As Long
    Dim tw As Long
    Dim fname As String
    
    On Error GoTo Err_Handler
    Me.picSave.Cls
    With BC
        .BarcodeOutput = Me.picSave
        .BarCodeX = 1
        .BarCodeY = 1
        .BarXFactor = 1
        .BarMultiplier = 1 'CInt(Me.cboMultiplier.Text)
        .BarcodeHeight = CLng(Me.txtHeight.Text) * (56.7 / Screen.TwipsPerPixelY)
        .BarTextGap = CInt(Me.txtGap.Text) * (56.7 / Screen.TwipsPerPixelY)
        .BarTextFont fname, fsize, fbold, fitalic
        .BarCodeData = txtBarcode.Text
        th = .TotalBarHeight
        tw = .TotalBarWidth
        Me.picSave.Width = (tw + 2) * Screen.TwipsPerPixelX
        Me.picSave.Height = (th + 2) * Screen.TwipsPerPixelY
        'Handle (hdc) has changed with the resize of the picture box so update it with the .dll
        .BarcodeOutput = Me.picSave
        .DrawBarCode
    End With
    Me.picSave.Refresh
    Me.picSave.Visible = True
    Me.lblImage.Top = Me.picSave.Top + Me.picSave.Height + 5
    Me.lblImage.Left = ((Me.picSave.Width - Me.lblImage.Width) / 2) + Me.picSave.Left
    Me.lblImage.Visible = True
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Save Barcode Image"
        .Filter = "Bitmap | *.bmp"
        .FileName = ""
        .ShowSave
        fname = Trim(.FileName)
    End With
    SavePicture Me.picSave.Image, fname
    Me.picSave.Visible = False
    Me.lblImage.Visible = False
    Exit Sub
    
    
Err_Handler:
    'Cancel error
    If Err.Number = 32755 Then
        Me.picSave.Visible = False
        Me.lblImage.Visible = False
        Exit Sub
    End If
    
    MsgBox "Error Number: " & CStr(Err.Number) & vbCrLf _
                        & "Description: " & Err.Description & vbCrLf _
                        & "Error Source: " & Err.Source, vbOKOnly + vbCritical, "Error"
End Sub

Private Sub Form_Load()
    Dim X As Integer
    
    Set BC = New clsLinearBarCodes
    Me.picWorkspace.ScaleMode = vbPixels
    Me.picSave.ScaleMode = vbPixels
    Me.lblHeading.Caption = "Select a Barcode"
    For X = 1 To 10
        Me.cboMultiplier.AddItem CStr(X), X - 1
    Next X
    Me.cboMultiplier.ListIndex = 0

    Me.cboXFactor.AddItem ".005", 0
    Me.cboXFactor.AddItem ".0075", 1
    Me.cboXFactor.ListIndex = 0
    
    Me.cboBarRatio.AddItem "3:1", 0
    Me.cboBarRatio.AddItem "2:1", 1
    Me.cboBarRatio.ListIndex = 0
    Me.mnuFilePrint.Enabled = False
    Me.mnuFileSaveBMP.Enabled = False
    Me.Show
    Me.txtBarcode.SetFocus
    CheckFields
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set BC = Nothing
End Sub

Private Sub mnuCode11_Click()
    Me.lblHeading.Caption = "Code 11"
    Me.chkCD.Enabled = True
    Me.chkDSS.Enabled = False
    Me.cboBarRatio.Enabled = False
    BC.Symbology = Code11
    CheckFields
End Sub

Private Sub mnuCode128a_Click()
    Me.lblHeading.Caption = "Code 128A"
    Me.chkDSS.Enabled = False
    Me.chkCD.Value = Unchecked
    Me.chkCD.Enabled = False
    Me.cboBarRatio.Enabled = False
    BC.Symbology = Code128_A
    CheckFields
End Sub

Private Sub mnuCode128Auto_Click()
    Me.lblHeading.Caption = "Code 128-Auto"
    Me.chkDSS.Enabled = False
    Me.chkCD.Value = Unchecked
    Me.chkCD.Enabled = False
    Me.cboBarRatio.Enabled = False
    BC.Symbology = Code128_Auto
    CheckFields
End Sub

Private Sub mnuCode128b_Click()
    Me.lblHeading.Caption = "Code 128B"
    Me.chkDSS.Enabled = False
    Me.chkCD.Value = Unchecked
    Me.chkCD.Enabled = False
    Me.cboBarRatio.Enabled = False
    BC.Symbology = Code128_B
    CheckFields
End Sub

Private Sub mnuCode128c_Click()
    Me.lblHeading.Caption = "Code 128C"
    Me.chkDSS.Enabled = False
    Me.chkCD.Value = Unchecked
    Me.chkCD.Enabled = False
    Me.cboBarRatio.Enabled = False
    BC.Symbology = Code128_C
    CheckFields
End Sub

Private Sub mnuCode2of5_Click()
    Me.lblHeading.Caption = "Code 2of5"
    Me.chkDSS.Enabled = False
    Me.chkCD.Value = Unchecked
    Me.chkCD.Enabled = False
    Me.cboBarRatio.Enabled = False
    BC.Symbology = Code2of5
    CheckFields
End Sub

Private Sub mnuCode39_Click()
    Me.lblHeading.Caption = "Code 39"
    Me.chkCD.Enabled = True
    Me.chkDSS.Enabled = True
    Me.cboBarRatio.Enabled = True
    BC.Symbology = Code39
    CheckFields
End Sub

Private Sub mnuCode39Extd_Click()
    Me.lblHeading.Caption = "Code 39 Full Ascii"
    Me.chkCD.Value = Unchecked
    Me.chkCD.Enabled = False
    Me.chkDSS.Enabled = True
    Me.cboBarRatio.Enabled = True
    BC.Symbology = Code39_Extd
    CheckFields
End Sub

Private Sub mnuCode93_Click()
    Me.lblHeading.Caption = "Code 93"
    Me.chkDSS.Enabled = True
    Me.chkCD.Value = Unchecked
    Me.chkCD.Enabled = False
    Me.cboBarRatio.Enabled = False
    BC.Symbology = Code93
    CheckFields
End Sub

Private Sub mnuCodeCodabar_Click()
    Me.lblHeading.Caption = "Codabar"
    Me.chkDSS.Enabled = False
    Me.chkCD.Value = Unchecked
    Me.chkCD.Enabled = False
    Me.cboBarRatio.Enabled = False
    BC.Symbology = Codabar
    CheckFields
End Sub

Private Sub mnuCodeI2of5_Click()
    Me.lblHeading.Caption = "Interleave 2of5"
    Me.chkDSS.Enabled = False
    Me.chkCD.Value = Unchecked
    Me.chkCD.Enabled = True
    Me.cboBarRatio.Enabled = True
    BC.Symbology = CodeI2of5
    CheckFields
End Sub

Private Sub mnuCodeMSI_Click()
    Me.lblHeading.Caption = "MSI"
    Me.chkDSS.Enabled = False
    Me.chkCD.Value = Unchecked
    Me.chkCD.Enabled = False
    Me.cboBarRatio.Enabled = False
    BC.Symbology = MSI
    CheckFields
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End
End Sub



Private Sub optHR_Click(Index As Integer)
    BC.HRText = Index
    cmdDraw_Click
End Sub

Private Sub optRotation_Click(Index As Integer)
Select Case Index
    Case 0
        BC.BarRotation = 0
    Case 1
        BC.BarRotation = 90
    Case 2
        BC.BarRotation = 270
    Case 3
        BC.BarRotation = 180
End Select
cmdDraw_Click
End Sub

Private Sub optTxtAlign_Click(Index As Integer)
    BC.HRTextAlignment = Index
    cmdDraw_Click
End Sub

Private Sub chkCD_Click()
    Me.chkDCD.Enabled = Me.chkCD.Value = Checked
    BC.UseCheckDigit = Me.chkCD.Value = Checked
End Sub

Private Sub chkDSS_Click()
    BC.DisplayStopStart = Me.chkDSS.Value = Checked
    cmdDraw_Click
End Sub

Private Sub chkDCD_Click()
    BC.DisplayCheckDigit = Me.chkDCD.Value = Checked
    cmdDraw_Click
End Sub
Private Sub optTxtPos_Click(Index As Integer)
If Index = 0 Then
    BC.HRTextPlacement = TP_UNDER
Else
    BC.HRTextPlacement = TP_ABOVE
End If
cmdDraw_Click
End Sub

Private Sub txtBarcode_Change()
    CheckFields
End Sub

Private Sub txtHeight_Change()
    CheckFields
End Sub

Private Sub txtX_Change()
  CheckFields
End Sub

Private Sub txtY_Change()
    CheckFields
End Sub


Private Sub CheckFields()
Dim sym As Integer

sym = BC.Symbology
If Me.txtX.Text = "" Or Me.txtY = "" Or Me.txtHeight = "" Or Me.txtBarcode = "" Or sym = 0 Then
        Me.cmdDraw.Enabled = False
        Me.cmdPrint.Enabled = False
        Me.mnuFilePrint.Enabled = False
        Me.mnuFileSaveBMP.Enabled = False
    Else
        Me.cmdDraw.Enabled = True
        Me.cmdPrint.Enabled = True
        Me.mnuFilePrint.Enabled = True
        Me.mnuFileSaveBMP.Enabled = True
    End If
End Sub

