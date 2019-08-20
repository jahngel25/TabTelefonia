VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReporte 
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   Icon            =   "frmReporte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   11100
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7155
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte.frx":014A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte.frx":046A
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte.frx":078A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte.frx":08E6
            Key             =   "Copy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbView 
      Height          =   3000
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   5292
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmReporte.frx":0A42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InitForm(psList As MSFlexGrid, psDescripcion As String)

   Dim i As Long, j As Integer, iStart As Long, sRow As String
   On Error GoTo ErrorManager
   ScaleMode = vbPixels
   rtbView.Text = psDescripcion
   rtbView.SelStart = 0
   rtbView.SelLength = Len(rtbView.Text)
   rtbView.SelFontSize = 9
   rtbView.SelBold = True
   rtbView.Text = rtbView.Text & vbCrLf & Now & vbCrLf
   iStart = Len(rtbView.Text)
   For i = 0 To psList.Rows - 1
        sRow = ""
        For j = 0 To psList.Cols - 1
            If Len(psList.TextMatrix(i, j)) > 25 Then
                sRow = sRow & Left(psList.TextMatrix(i, j), 25) & vbTab
            Else
                sRow = sRow & IIf(Len(psList.TextMatrix(i, j)) < 10, Left(psList.TextMatrix(i, j) & "                          ", 10), psList.TextMatrix(i, j)) & vbTab
            End If
        Next
        rtbView.Text = rtbView.Text & vbCrLf & sRow
        If i = 0 Then
            rtbView.SelStart = iStart
            rtbView.SelLength = Len(rtbView.Text)
            rtbView.SelUnderline = True
            iStart = iStart + Len(sRow) + 1
        End If
   Next i
    rtbView.SelStart = iStart + 1
    rtbView.SelLength = Len(rtbView.Text)
    rtbView.SelBold = False
    rtbView.SelTabCount = 5
    rtbView.SelTabs(0) = 70
    rtbView.SelTabs(1) = 250
    rtbView.SelTabs(2) = 100
    rtbView.SelTabs(3) = 240
    rtbView.SelTabs(4) = 250
    rtbView.SelBold = False
    rtbView.SelColor = vbBlack
    rtbView.SelFontSize = 8
   Exit Sub
    
    
ErrorManager:
    SubGMuestraError

End Sub

Private Sub Form_Resize()

   With rtbView
      .Left = 1
      .Top = Toolbar1.Height + 10
      .Width = Me.Width
      .Height = Me.Height - Toolbar1.Height + 10
   End With

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   
   Select Case Button.Key
      Case "Exit"
         Unload Me
      Case "Print"
         rtbView.SelStart = 0
         rtbView.SelLength = Len(rtbView.Text)
         Printer.Font = rtbView.Font.Name
         Printer.FontSize = rtbView.Font.Size
         rtbView.SelPrint Printer.hDC
      Case "Cut"
         rtbView.SelStart = 0
         rtbView.SelLength = Len(rtbView.Text)
         Clipboard.SetText rtbView.SelText
         rtbView.SelText = ""
      Case "Copy"
         rtbView.SelStart = 0
         rtbView.SelLength = Len(rtbView.Text)
         Clipboard.SetText rtbView.SelText
   End Select
    
End Sub
