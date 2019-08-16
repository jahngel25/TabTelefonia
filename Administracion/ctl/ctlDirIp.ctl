VERSION 5.00
Begin VB.UserControl EditIPBox 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ScaleHeight     =   300
   ScaleWidth      =   2175
   Begin VB.ComboBox cmbIP 
      Height          =   315
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   0
      Width           =   675
   End
   Begin VB.TextBox txt_ip 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   0
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "255"
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox txt_ip 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   495
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "255"
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox txt_ip 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   990
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "255"
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox txt_ip 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "255"
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox txt_ipdot 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   345
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   " ."
      Top             =   0
      Width           =   120
   End
   Begin VB.TextBox txt_ipdot 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   " ."
      Top             =   0
      Width           =   120
   End
   Begin VB.TextBox txt_ipdot 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1335
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   " ."
      Top             =   0
      Width           =   120
   End
End
Attribute VB_Name = "EditIPBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : EditIPBox
' Fecha  : 24/09/2004 08:09
' Author    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   : Permitir el ingreso y validación de direcciones IP
'---------------------------------------------------------------------------------------
Option Explicit
Dim IPvalidated  As Boolean
Dim Cnt As Integer
Dim IPaddress As String
Dim bTieneFoco As Boolean
Dim bColocarFoco As Boolean
Dim bMostrarCombo As Integer
Public Event Change()
Public Event ComboClick()
Public Event DoLostFocus()
Public Event SetFocus()

Private Sub Clear()
   On Error GoTo ErrorManager

    With txt_ip
        For Cnt = 0 To .UBound
            .Item(Cnt).Text = ""
            .Item(Cnt).Tag = "False"
        Next Cnt
    End With

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Validate()
   On Error GoTo ErrorManager

    IPvalidated = True
    IPaddress = ""
    
    With txt_ip
        For Cnt = 0 To .UBound
            If .Item(Cnt).Tag = "False" Then
                IPvalidated = False
                '.Item(Cnt).SetFocus
                txt_ip_GotFocus (Cnt)
                Exit For
            End If
        Next Cnt
    If IPvalidated = False Then
        'MsgBox ("No es una dirección IP Válida")
    Else
        For Cnt = 0 To txt_ip.UBound
            IPaddress = IPaddress & Trim(.Item(Cnt).Text) & "."
        Next Cnt
        IPaddress = Left(IPaddress, Len(IPaddress) - 1)
    End If
    End With

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmbIP_Click()
   On Error GoTo ErrorManager

    txt_ip(3).Text = cmbIP.List(cmbIP.ListIndex)
    txt_ip.Item(3).Tag = True
    RaiseEvent Change
    
   RaiseEvent ComboClick

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmbIP_LostFocus()
   On Error GoTo ErrorManager

    RaiseEvent DoLostFocus

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub txt_ip_Change(Index As Integer)
Dim Chr_Cnt As Integer
Dim New_Txt As String
   On Error GoTo ErrorManager

    If bTieneFoco = True Then
        bTieneFoco = False
        Exit Sub
    End If
    With txt_ip(Index)
        If Len(.Text) = 0 Then Exit Sub
        For Chr_Cnt = 1 To Len(.Text)
            If Mid(.Text, Chr_Cnt, 1) = "." Or Mid(.Text, Chr_Cnt, 1) = " " Then
                If Index < 3 Then
                    'txt_ip(Index + 1).SetFocus
                End If
            End If
            If Mid(.Text, Chr_Cnt, 1) >= "0" And Mid(.Text, Chr_Cnt, 1) <= "9" Then
                New_Txt = New_Txt & Mid(.Text, Chr_Cnt, 1)
            End If
        Next Chr_Cnt
        .Text = New_Txt
        If bColocarFoco And Len(.Text) = 3 And Index < 3 Then
            If txt_ip(Index + 1).Visible Then
                txt_ip(Index + 1).SetFocus
            Else
                If Index = 2 Then cmbIP.SetFocus
            End If
        End If
        If Val(.Text) > 255 Then
            .ForeColor = &HC0&
            .Tag = "False"
        Else
            .ForeColor = &H80000008
            .Tag = "True"
        End If
    End With
    RaiseEvent Change

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub txt_ip_GotFocus(Index As Integer)
   On Error GoTo ErrorManager

    txt_ip(Index).SelStart = 0
    txt_ip(Index).SelLength = Len(txt_ip(Index).Text)
    txt_ip(Index).Text = LTrim(txt_ip(Index).Text)
    bTieneFoco = False

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub txt_ip_KeyPress(Index As Integer, KeyAscii As Integer)
        
   On Error GoTo ErrorManager

        If KeyAscii = 13 Then
            If Index = 2 And bMostrarCombo Then
                cmbIP.SetFocus
            Else
                SendKeys "{Tab}"
                KeyAscii = 0
            End If
        End If

      Exit Sub
ErrorManager:
    SubGMuestraError
        
End Sub

Private Sub txt_ip_LostFocus(Index As Integer)
   On Error GoTo ErrorManager

    bTieneFoco = True
    With txt_ip(Index)
        If .Text <> "" Then
            .Text = Val(.Text)
            bTieneFoco = True
        Else
            .Tag = "False"
        End If
    End With

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Public Property Let TextItem(ByVal Index As Integer, ByVal new_value As String)
    On Error GoTo ErrorManager
    bColocarFoco = False
    txt_ip(Index) = new_value
    bColocarFoco = True
Exit Property
ErrorManager:
SubGMuestraError
End Property

Public Property Let MostrarCombo(bValor As Boolean)
   On Error GoTo ErrorManager

    bMostrarCombo = bValor
    txt_ip.Item(3).Visible = Not bValor
Exit Property
ErrorManager:
    SubGMuestraError
End Property

Public Property Let ItemEnabled(ByVal Index As Integer, ByVal new_value As String)
    On Error GoTo ErrorManager
    
    txt_ip(Index).Enabled = new_value
    
Exit Property
ErrorManager:
SubGMuestraError
End Property

Public Property Get TextItem(ByVal Index As Integer) As String
    On Error GoTo ErrorManager
    TextItem = txt_ip(Index)
Exit Property
ErrorManager:
SubGMuestraError
End Property

Public Property Get Text() As String
    On Error GoTo ErrorManager
    Call Validate
    Text = IPaddress

Exit Property
ErrorManager:
SubGMuestraError
End Property

Private Sub UserControl_GotFocus()
   On Error GoTo ErrorManager

       txt_ip_GotFocus (0)

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub UserControl_Initialize()
   On Error GoTo ErrorManager

    Call Clear
    bColocarFoco = True

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Public Sub LlenarCombo(iComienzo As Integer, iFinal As Integer, iAncho As Integer, bMostrarPrimero As Boolean)
    Dim i As Integer
   On Error GoTo ErrorManager

    cmbIP.Clear
    For i = iComienzo To iFinal Step iAncho
        If (i = iComienzo) Then
            If bMostrarPrimero Then cmbIP.AddItem i
        Else
            cmbIP.AddItem i
        End If
    Next

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

