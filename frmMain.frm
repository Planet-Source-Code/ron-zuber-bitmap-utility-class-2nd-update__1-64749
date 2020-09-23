VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMP Utils Example Application"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtExt 
      Height          =   300
      Left            =   3720
      TabIndex        =   0
      Top             =   5880
      Width           =   750
   End
   Begin VB.CommandButton cmdRevert 
      Caption         =   "Select/Revert BMP(s)"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4635
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Select the BMP file(s) to convert, and then show a test view."
      Top             =   5760
      Width           =   1875
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Converted File"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4635
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Select the BMP file(s) to convert, and then show a test view."
      Top             =   5280
      Width           =   1875
   End
   Begin MSComctlLib.ProgressBar pbCurr 
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   6240
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Select/Convert BMP(s)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   330
      Left            =   4635
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Select the BMP file(s) to convert, and then show a test view."
      Top             =   4785
      Width           =   1875
   End
   Begin MSComctlLib.ProgressBar pbAll 
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   6525
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox picViewBack 
      AutoRedraw      =   -1  'True
      Height          =   4740
      Left            =   0
      ScaleHeight     =   312
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   430
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   6510
      Begin VB.HScrollBar hsHorizontal 
         Height          =   225
         LargeChange     =   100
         Left            =   -15
         Max             =   10
         SmallChange     =   10
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4455
         Visible         =   0   'False
         Width           =   6270
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6255
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   12
         Top             =   4485
         Width           =   210
      End
      Begin VB.VScrollBar vsVertical 
         Height          =   4485
         LargeChange     =   1000
         Left            =   6225
         Max             =   100
         SmallChange     =   10
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.PictureBox picView 
         AutoRedraw      =   -1  'True
         Height          =   150
         Left            =   -15
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   -15
         Visible         =   0   'False
         Width           =   150
      End
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Options"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   4455
      Begin VB.CheckBox chkEncrypt 
         Caption         =   "Encryption"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtKey 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   19
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton oZLib 
         Caption         =   "ZLib"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton oNTNative 
         Caption         =   "NT Native"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox chkCompression 
         Caption         =   "Compression"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Key"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   960
      End
   End
   Begin VB.Label lblWhat 
      Height          =   225
      Left            =   0
      TabIndex        =   7
      Top             =   7080
      Width           =   6495
   End
   Begin VB.Label lblFile 
      Height          =   225
      Left            =   0
      TabIndex        =   5
      Top             =   6840
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Choose a file extension for the converted BMP(s):"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   5955
      Width           =   3510
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Processing      As Boolean
Private Sub ShowScrollBars()
  hsHorizontal.Value = 0
  vsVertical.Value = 0
  If picView.Width > picViewBack.Width - 16 Then
    hsHorizontal.Visible = True
    hsHorizontal.Max = picView.Width - picViewBack.Width + 16
    hsHorizontal.SmallChange = 1
    hsHorizontal.LargeChange = picViewBack.Width * 0.75
  Else
    hsHorizontal.Visible = False
  End If
  If picView.Height > picViewBack.Height - 16 Then
    vsVertical.Visible = True
    vsVertical.Max = picView.Height - picViewBack.Height + 16
    vsVertical.SmallChange = 1
    vsVertical.LargeChange = picViewBack.Height * 0.75
  Else
    vsVertical.Visible = False
  End If
End Sub
Private Sub chkCompression_Click()
  If chkCompression.Value = 0 Then
    oNTNative.Enabled = False
    oZLib.Enabled = False
  Else
    oNTNative.Enabled = True
    oZLib.Enabled = True
  End If
End Sub
Private Sub chkEncrypt_Click()
  If chkEncrypt.Value = 0 Then
    txtKey.Enabled = False
    txtKey.BackColor = Me.BackColor
  Else
    txtKey.Enabled = True
    txtKey.BackColor = &H80000005
  End If
End Sub
Private Sub hsHorizontal_Change()
  picView.Left = -1 + hsHorizontal.Value * -1
End Sub
Private Sub hsHorizontal_Scroll()
  picView.Left = -1 + hsHorizontal.Value * -1
End Sub
Private Sub oNTNative_Click()
  If oNTNative.Value = True Then oZLib.Value = False
End Sub
Private Sub oZLib_Click()
  If oZLib.Value = True Then oNTNative.Value = False
End Sub
Private Sub vsVertical_Change()
  picView.Top = -1 + vsVertical.Value * -1
End Sub
Private Sub vsVertical_Scroll()
  picView.Top = -1 + vsVertical.Value * -1
End Sub
Private Sub cmdLoad_Click()
  Dim Key As String
  Dim BMC As BitmapUtils
  
  'If we're using encryption, check to see if a key was entered.
  If chkEncrypt.Value <> 0 And Len(txtKey.Text) = 0 Then
    Call MsgBox("You must enter a key to utilize encryption.", vbExclamation, "Zero Length Key")
    Exit Sub
  End If
  
  On Error Resume Next
  With cDialog
    .CancelError = True
    .DefaultExt = txtExt.Text
    .DialogTitle = "Select a converted file."
    .FileName = ""
    .Filter = "Converted Bitmap (*." & txtExt.Text & ")|*." & txtExt.Text
    .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNExplorer
    .MaxFileSize = 256
    .ShowOpen
  End With
  
  If Err.Number <> 0 Then
    If Err.Number = cdlCancel Then
      txtExt.Enabled = True
      cmdConvert.Enabled = True
      cmdLoad.Enabled = True
      Err.Clear
      Exit Sub
    End If
    Call MsgBox("Run-Time Error (" & Err.Number & "): " & Err.Description, vbOKOnly, "Run-Time Error " & Err.Number)
    End
  End If
  On Error GoTo 0
  
  lblFile.Caption = "Current File: " & Right(cDialog.FileName, Len(cDialog.FileName) - InStrRev(cDialog.FileName, Chr(0)))
  lblWhat.Caption = "What: Loading File..."
  DoEvents
  
  Set BMC = New BitmapUtils
    
  Call BMC.LoadByteData(cDialog.FileName)
  
  If chkEncrypt.Value <> 0 Then
    lblWhat.Caption = "What: Decrypting file..."
    DoEvents
    
    Key = txtKey.Text
  
    Call BMC.DecryptByteData(Key)
  End If
  
  If chkCompression.Value <> 0 Then
    lblWhat.Caption = "What: Decompressing file..."
    DoEvents
    
    If oNTNative.Value = True Then
      Call BMC.DecompressByteData
    Else
      Call BMC.DecompressByteData_ZLib
    End If
  End If
  
  picView.Cls
  picView.Height = BMC.ImageHeight + 4
  picView.Width = BMC.ImageWidth + 4
  picView.Visible = True
  ShowScrollBars
  Call BMC.Blt(picView.hDC)
  picView.Refresh
  Set BMC = Nothing
End Sub
Private Sub cmdConvert_Click()
  Dim BMC As BitmapUtils
  Dim i As Long, j As Long
  Dim Files() As String
  Dim Key As String
  
  'If we're using encryption, check to see if a key was entered.
  If chkEncrypt.Value <> 0 And Len(txtKey.Text) = 0 Then
    Call MsgBox("You must enter a key to utilize encryption.", vbExclamation, "Zero Length Key")
    Exit Sub
  End If
  
  On Error Resume Next
  With cDialog
    .CancelError = True
    .DefaultExt = "BMP"
    .DialogTitle = "Select Bitmap File(s)"
    .FileName = ""
    .Filter = "Bitmap Files (*.bmp)|*.bmp"
    .Flags = cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNExplorer
    .MaxFileSize = 32000
    .ShowOpen
  End With
  
  If Err.Number <> 0 Then
    If Err.Number = cdlCancel Then
      Err.Clear
      Exit Sub
    End If
    Call MsgBox("Run-Time Error (" & Err.Number & "): " & Err.Description, vbOKOnly, "Run-Time Error " & Err.Number)
    End
  End If
  On Error GoTo 0
  
  Processing = True 'Let the rest of the application know we're processing files.
  txtExt.Enabled = False
  txtKey.Enabled = False
  cmdConvert.Enabled = False
  cmdLoad.Enabled = False
  cmdRevert.Enabled = False
  
  If InStr(cDialog.FileName, Chr(0)) Then
    Files = Split(cDialog.FileName, Chr(0))
  Else
    ReDim Files(1)
    i = InStrRev(cDialog.FileName, "\")
    Files(0) = Left$(cDialog.FileName, i - 1)
    Files(1) = Right$(cDialog.FileName, Len(cDialog.FileName) - i)
  End If
  
  pbCurr.Max = 2
  
  If chkCompression.Value <> 0 Then
    pbCurr.Max = pbCurr.Max + 1
  End If
  If chkEncrypt.Value <> 0 Then
    pbCurr.Max = pbCurr.Max + 1
  End If
  
  pbAll.Max = UBound(Files) * pbCurr.Max
  pbAll.Value = 0
  
  For i = 1 To UBound(Files)
    Set BMC = New BitmapUtils
    pbCurr.Value = 0
    lblFile.Caption = "Current File: " & Files(i) & "..."
    lblWhat.Caption = "What: Loading File..."
    DoEvents
    
    Call BMC.LoadByteData(Files(0) & "\" & Files(i))
    
    picView.Cls
    picView.Height = BMC.ImageHeight + 4
    picView.Width = BMC.ImageWidth + 4
    picView.Visible = True
    ShowScrollBars
    Call BMC.Blt(picView.hDC)
    
    If chkCompression.Value <> 0 Then
      pbCurr.Value = pbCurr.Value + 1
      pbAll.Value = pbAll.Value + 1
      lblWhat.Caption = "What: Compressing Bytes..."
      DoEvents
      
      If oNTNative.Value = True Then
        Call BMC.CompressByteData
      Else
        Call BMC.CompressByteData_ZLib
      End If
    End If
    
    If chkEncrypt.Value <> 0 Then
      pbCurr.Value = pbCurr.Value + 1
      pbAll.Value = pbAll.Value + 1
      lblWhat.Caption = "What: Encrypting Bytes..."
      DoEvents
      
      Key = txtKey.Text
      
      Call BMC.EncryptByteData(Key)
    End If
      
    pbCurr.Value = pbCurr.Value + 1
    pbAll.Value = pbAll.Value + 1
    lblWhat.Caption = "Saving File..."
    DoEvents
    
    Files(i) = Left$(Files(i), Len(Files(i)) - 3) & txtExt.Text
    
    If Not Dir(Files(0) & "\" & Files(i)) = "" Then
      Call Kill(Files(0) & "\" & Files(i))
    End If
    
    Call BMC.SaveByteData(Files(0) & "\" & Files(i))
    
    pbCurr.Value = pbCurr.Value + 1
    pbAll.Value = pbAll.Value + 1
    DoEvents
    
    Set BMC = Nothing
  Next
cmdConvert_Exit:
  Processing = False 'Let everyone know we're done. :)
  txtExt.Enabled = True
  txtKey.Enabled = True
  cmdConvert.Enabled = True
  cmdLoad.Enabled = True
  cmdRevert.Enabled = True
  Set BMC = Nothing
End Sub
Private Sub cmdRevert_Click()
  Dim BMC As BitmapUtils
  Dim i As Long, j As Long
  Dim Files() As String
  Dim Key As String
  
  With cDialog
    .CancelError = True
    .DefaultExt = "BMP"
    .DialogTitle = "Select Bitmap File(s)"
    .FileName = ""
    .Filter = "Converted Bitmap (*." & txtExt.Text & ")|*." & txtExt.Text
    .Flags = cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNExplorer
    .MaxFileSize = 32000
    .ShowOpen
  End With
  
  On Error Resume Next
  If Err.Number <> 0 Then
    If Err.Number = cdlCancel Then
      Err.Clear
      Exit Sub
    End If
    Call MsgBox("Run-Time Error (" & Err.Number & "): " & Err.Description, vbOKOnly, "Run-Time Error " & Err.Number)
    End
  End If
  On Error GoTo 0
  
  Processing = True 'Let the rest of the application know we're processing files.
  txtExt.Enabled = False
  txtKey.Enabled = False
  cmdConvert.Enabled = False
  cmdLoad.Enabled = False
  cmdRevert.Enabled = False
  
  If InStr(cDialog.FileName, Chr(0)) Then
    Files = Split(cDialog.FileName, Chr(0))
  Else
    ReDim Files(1)
    i = InStrRev(cDialog.FileName, "\")
    Files(0) = Left$(cDialog.FileName, i - 1)
    Files(1) = Right$(cDialog.FileName, Len(cDialog.FileName) - i)
  End If
  
  pbCurr.Max = 2
  
  If chkCompression.Value <> 0 Then
    pbCurr.Max = pbCurr.Max + 1
  End If
  If chkEncrypt.Value <> 0 Then
    pbCurr.Max = pbCurr.Max + 1
  End If
  
  pbAll.Max = UBound(Files) * pbCurr.Max
  pbAll.Value = 0
  
  For i = 1 To UBound(Files)
    Set BMC = New VWBitmapUtils.BitmapUtils
    pbCurr.Value = 0
    lblFile.Caption = "Current File: " & Files(i) & "..."
    lblWhat.Caption = "What: Loading File..."
    DoEvents
    
    On Error Resume Next
    Call BMC.LoadByteData(Files(0) & "\" & Files(i))
    'Call BMC.SaveBitmap("C:\Test.bmp")
    If Err.Number <> 0 Then
      Call MsgBox("Run-Time Error (" & Err.Number & "): " & Err.Description, vbOKOnly, "Error While Loading Bitmap")
      Err.Clear
      GoTo cmdRevert_Exit
    End If
    On Error GoTo 0
    
    If chkEncrypt.Value <> 0 Then
      pbCurr.Value = pbCurr.Value + 1
      pbAll.Value = pbAll.Value + 1
      lblWhat.Caption = "What: Decrypting Bytes..."
      DoEvents
      
      Key = txtKey.Text
      
      Call BMC.DecryptByteData(Key)
    End If
    
    If chkCompression.Value <> 0 Then
      pbCurr.Value = pbCurr.Value + 1
      pbAll.Value = pbAll.Value + 1
      lblWhat.Caption = "What: Decompressing Bytes..."
      DoEvents
      
      If oNTNative.Value = True Then
        Call BMC.DecompressByteData
      Else
        Call BMC.DecompressByteData_ZLib
      End If
    End If
    
    pbCurr.Value = pbCurr.Value + 1
    pbAll.Value = pbAll.Value + 1
    lblWhat.Caption = "Saving File..."
    DoEvents
    
    Files(i) = Left$(Files(i), Len(Files(i)) - 3) & "BMP"
    
    If Not Dir(Files(0) & "\" & Files(i)) = "" Then
      Call Kill(Files(0) & "\" & Files(i))
    End If
    
    picView.Cls
    picView.Height = BMC.ImageHeight + 4
    picView.Width = BMC.ImageWidth + 4
    picView.Visible = True
    ShowScrollBars
    Call BMC.Blt(picView.hDC)
    DoEvents
    
    Call BMC.SaveByteData(Files(0) & "\" & Files(i))
    
    pbCurr.Value = pbCurr.Value + 1
    pbAll.Value = pbAll.Value + 1
    DoEvents
    
    Set BMC = Nothing
  Next
cmdRevert_Exit:
  Processing = False 'Let everyone know we're done. :)
  txtExt.Enabled = True
  txtKey.Enabled = True
  cmdConvert.Enabled = True
  cmdLoad.Enabled = True
  cmdRevert.Enabled = True
  Set BMC = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If Processing Then
    If MsgBox("You are currently processing files. If you quit now, your progress on unsaved files will be lost. Do you still wish to quit?", vbYesNo, "Currently Processing Files") = vbNo Then
      Cancel = True
    Else
      End
    End If
  End If
End Sub
Private Sub txtExt_Change()
  If Len(txtExt.Text) = 0 Then
    cmdConvert.Enabled = False
    cmdLoad.Enabled = False
    cmdRevert.Enabled = False
  Else
    cmdConvert.Enabled = True
    cmdLoad.Enabled = True
    cmdRevert.Enabled = True
  End If
End Sub
