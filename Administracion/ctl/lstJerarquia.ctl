VERSION 5.00
Begin VB.UserControl ctlLstJerarquia 
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   ScaleHeight     =   1890
   ScaleWidth      =   2550
   Begin VB.CommandButton cmdEliminar 
      Height          =   225
      Left            =   2295
      Picture         =   "lstJerarquia.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Eliminar"
      Top             =   45
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdAdicionar 
      Height          =   225
      Left            =   2025
      Picture         =   "lstJerarquia.ctx":024A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Adicionar"
      Top             =   45
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdAbajo 
      Height          =   240
      Left            =   30
      Picture         =   "lstJerarquia.ctx":0494
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Bajar"
      Top             =   1590
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdOrdenarAZ 
      Height          =   645
      Left            =   30
      Picture         =   "lstJerarquia.ctx":0686
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ordenar"
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdArriba 
      Height          =   240
      Left            =   30
      Picture         =   "lstJerarquia.ctx":0BD8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Subir"
      Top             =   270
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ListBox lstDescripcion 
      Height          =   1620
      ItemData        =   "lstJerarquia.ctx":0DCA
      Left            =   270
      List            =   "lstJerarquia.ctx":0DCC
      TabIndex        =   0
      Top             =   240
      Width           =   2265
   End
   Begin VB.ListBox lstCodigo 
      Height          =   1620
      ItemData        =   "lstJerarquia.ctx":0DCE
      Left            =   270
      List            =   "lstJerarquia.ctx":0DD0
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.ListBox lstAdicional 
      Height          =   255
      ItemData        =   "lstJerarquia.ctx":0DD2
      Left            =   2070
      List            =   "lstJerarquia.ctx":0DD9
      TabIndex        =   8
      Top             =   1380
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTexto 
      Caption         =   "Nombre del campo"
      Height          =   225
      Left            =   15
      TabIndex        =   7
      Top             =   15
      Width           =   1980
   End
End
Attribute VB_Name = "ctlLstJerarquia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : ctlLstJerarquia
' Fecha  : 24/09/2004 08:09
' Author    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   : Lista enriquecida para permitir ordenamiento,movimiento de los items, botón adicionar y botón eliminar
'---------------------------------------------------------------------------------------

Option Explicit
Private WithEvents TheFont As StdFont
Attribute TheFont.VB_VarHelpID = -1
Private bMostrarBotonesH As Boolean 'Horizontal
Private bMostrarBotonesV As Boolean 'Vertical
Private bEnabled As Boolean
Public Event Click()
Public Event DblClick()
Public Event btnEliminarClick()
Public Event btnAdicionarClick()
Public Event ListLostFocus()
Public Event ListGotFocus()
Public Event CambioOrden()
Public Event CambioOrdenAZ()

Private Sub cmdAbajo_Click()

On Error GoTo ErrorManager

    Dim strTemp1 As String
    Dim iCnt    As Integer
    Dim tmpCodigo As String
    Dim tmpAdicional As String
    iCnt = lstDescripcion.ListIndex
    
    If iCnt > -1 And iCnt < Me.ListCount - 1 Then
        tmpCodigo = lstCodigo.List(iCnt)
        strTemp1 = lstDescripcion.List(iCnt)
        tmpAdicional = lstAdicional.List(iCnt)
        AddItem strTemp1, tmpCodigo, tmpAdicional, (iCnt + 2)
        RemoveItem (iCnt)
        lstDescripcion.Selected(iCnt + 1) = True
   End If
    RaiseEvent CambioOrden
    
Exit Sub
ErrorManager:
    SubGMuestraError
    
End Sub

Private Sub cmdAdicionar_Click()
On Error GoTo ErrorManager
    
    RaiseEvent btnAdicionarClick
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdArriba_Click()
 On Error GoTo ErrorManager
 
    Dim strTemp1 As String
    Dim iCnt    As Integer
    Dim tmpCodigo As String
    Dim tmpAdicional As String
    
    iCnt = lstDescripcion.ListIndex
    If iCnt > 0 Then
         strTemp1 = lstDescripcion.List(iCnt)
         tmpCodigo = lstCodigo.List(iCnt)
         tmpAdicional = lstAdicional.List(iCnt)
        AddItem strTemp1, tmpCodigo, tmpAdicional, (iCnt - 1)
        RemoveItem (iCnt + 1)
        lstDescripcion.Selected(iCnt - 1) = True
    End If
    RaiseEvent CambioOrden
    
        Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo ErrorManager
    
    RaiseEvent btnEliminarClick
    
Exit Sub
ErrorManager:
    SubGMuestraError

End Sub

Private Sub cmdOrdenarAZ_Click()
Dim i As Integer, i2 As Integer, Hold As String, HoldID As String, HoldAdicional As String
Dim ValPer1 As String, ValPer2 As String

On Error GoTo ErrorManager

For i = 0 To lstDescripcion.ListCount - 1
    For i2 = 0 To lstDescripcion.ListCount - 1
        If i <> i2 Then
            ValPer1 = lstDescripcion.List(i)
            ValPer2 = lstDescripcion.List(i2)
            If ValPer1 < ValPer2 Then
                Hold = lstDescripcion.List(i)
                HoldID = lstCodigo.List(i)
                HoldAdicional = lstAdicional.List(i)
                
                lstDescripcion.List(i) = lstDescripcion.List(i2)
                lstDescripcion.List(i2) = Hold
                
                lstCodigo.List(i) = lstCodigo.List(i2)
                lstCodigo.List(i2) = HoldID
                
                lstAdicional.List(i) = lstAdicional.List(i2)
                lstAdicional.List(i2) = HoldAdicional
            End If
        End If
    Next
Next
RaiseEvent CambioOrdenAZ

Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub UserControl_EnterFocus()
On Error GoTo ErrorManager
    
    RaiseEvent ListGotFocus
    cmdAdicionar.Visible = bMostrarBotonesH And bMostrarBotonesV
    cmdEliminar.Visible = bMostrarBotonesH And bMostrarBotonesV
    cmdArriba.Visible = bMostrarBotonesH And bMostrarBotonesV
    cmdOrdenarAZ.Visible = bMostrarBotonesH And bMostrarBotonesV
    cmdAbajo.Visible = bMostrarBotonesH And bMostrarBotonesV

Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub UserControl_ExitFocus()

On Error GoTo ErrorManager
    
    cmdAdicionar.Visible = False
    cmdEliminar.Visible = False
    cmdArriba.Visible = False
    cmdOrdenarAZ.Visible = False
    cmdAbajo.Visible = False
    RaiseEvent ListLostFocus
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub UserControl_Initialize()
On Error GoTo ErrorManager
    
    Call UserControl_Resize
    Set TheFont = New StdFont
    TheFont.Name = "Ms Sans Serif"
    TheFont.Size = 8
    bMostrarBotonesH = False
    bMostrarBotonesV = False
    
Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error GoTo ErrorManager

    ListIndex = PropBag.ReadProperty("ListIndex", -1)
    lstDescripcion.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    lstDescripcion.ForeColor = PropBag.ReadProperty("ForeColor", 0)
    Set FontInfo = PropBag.ReadProperty("FontInfo", UserControl.Font)
    bMostrarBotonesV = PropBag.ReadProperty("MostrarBotonesV", False)
    bMostrarBotonesH = PropBag.ReadProperty("MostrarBotonesH", False)
    bEnabled = PropBag.ReadProperty("Enabled", True)
Exit Sub
ErrorManager:
    SubGMuestraError
    
End Sub

Private Sub UserControl_Resize()
    
    On Error GoTo ErrorManager
    

    lstDescripcion.Width = Width - lstDescripcion.Left
    lstDescripcion.Height = Height - 240
    lstDescripcion.Top = 270
    lstDescripcion.Left = 270
    cmdEliminar.Top = 45
    cmdEliminar.Left = Width - cmdEliminar.Width
    cmdAdicionar.Left = cmdEliminar.Left - cmdAdicionar.Width
    cmdAdicionar.Top = 45
    cmdArriba.Top = lstDescripcion.Top
    cmdOrdenarAZ.Top = (Height - cmdOrdenarAZ.Height) / 2
    cmdAbajo.Top = lstDescripcion.Top + lstDescripcion.Height - cmdAbajo.Height
    cmdArriba.Left = 30
    cmdOrdenarAZ.Left = 30
    cmdAbajo.Left = 30
    
 Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Public Sub RemoveItem(Index As Integer)

On Error GoTo ErrorManager

    Dim i As Integer
    If Index = -1 Or Index > ListCount Then Exit Sub
    lstDescripcion.RemoveItem Index
    lstCodigo.RemoveItem Index
    lstAdicional.RemoveItem Index
    
 Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Public Sub AddItem(proDescription As String, proID As String, Optional sCampoAdicional As String, Optional Index As Integer = -1)
    
    On Error GoTo ErrorManager
    If Index = -1 Then
        lstDescripcion.AddItem proDescription
        lstCodigo.AddItem proID
        lstAdicional.AddItem sCampoAdicional
    Else
        lstDescripcion.AddItem proDescription, Index
        lstCodigo.AddItem proID, Index
        lstAdicional.AddItem sCampoAdicional, Index
    End If
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Public Function ListDescripcion(Index As Integer) As String
On Error GoTo ErrorManager

        
        ListDescripcion = lstDescripcion.List(Index)
        
Exit Function
ErrorManager:
    SubGMuestraError
End Function

Public Function ListCampoAdicional(Index As Integer) As String
On Error GoTo ErrorManager

        
        ListCampoAdicional = lstAdicional.List(Index)
        
Exit Function
ErrorManager:
    SubGMuestraError
End Function

Public Function ListCodigo(Index As Integer) As String
    
    On Error GoTo ErrorManager

        ListCodigo = Trim(lstCodigo.List(Index))
        
Exit Function
ErrorManager:
    SubGMuestraError
End Function

Public Sub Clear()
    
On Error GoTo ErrorManager

    lstDescripcion.Clear
    lstCodigo.Clear
    lstAdicional.Clear

Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub lstDescripcion_Click()
On Error GoTo ErrorManager

    RaiseEvent Click
    
Exit Sub
ErrorManager:
    SubGMuestraError
    
End Sub

Private Sub lstDescripcion_DblClick()
On Error GoTo ErrorManager
        
    RaiseEvent DblClick
    
Exit Sub
ErrorManager:
    SubGMuestraError

End Sub

Public Property Get ListIndex() As Integer
    
On Error GoTo ErrorManager

    ListIndex = lstDescripcion.ListIndex
    
Exit Property
ErrorManager:
    SubGMuestraError
End Property

Public Property Get ListForeColor() As OLE_COLOR
    
On Error GoTo ErrorManager

    ListForeColor = lstDescripcion.ForeColor

Exit Property
ErrorManager:
    SubGMuestraError
End Property

Public Property Let MostrarBotonesH(bMostrar As Boolean)
  On Error GoTo ErrorManager

    bMostrarBotonesH = bMostrar
    PropertyChanged "MostrarBotonesH"

Exit Property
ErrorManager:
    SubGMuestraError
       
End Property

Public Property Get MostrarBotonesH() As Boolean
    On Error GoTo ErrorManager

    MostrarBotonesH = bMostrarBotonesH

Exit Property
ErrorManager:
    SubGMuestraError
End Property

Public Property Let MostrarBotonesV(bMostrar As Boolean)
    On Error GoTo ErrorManager

    bMostrarBotonesV = bMostrar
    PropertyChanged "MostrarBotonesV"

Exit Property
ErrorManager:
    SubGMuestraError
End Property

Public Property Let Enabled(bEnabled As Boolean)
    On Error GoTo ErrorManager

    MostrarBotonesV = bEnabled
    lstDescripcion.Enabled = bEnabled
    PropertyChanged "Enabled"

Exit Property
ErrorManager:
    SubGMuestraError
End Property

Public Property Get Enabled() As Boolean
 On Error GoTo ErrorManager

    Enabled = bEnabled
    
Exit Property
ErrorManager:
    SubGMuestraError
End Property

Public Property Get MostrarBotonesV() As Boolean
    On Error GoTo ErrorManager
    
    MostrarBotonesV = bMostrarBotonesV

Exit Property
ErrorManager:
    SubGMuestraError
End Property

Public Property Let ListForeColor(ByVal proColor As OLE_COLOR)
    On Error GoTo ErrorManager

    lstDescripcion.ForeColor = proColor
    PropertyChanged "ListForeColor"


Exit Property
ErrorManager:
    SubGMuestraError
       
End Property

Public Property Get ListBackColor() As OLE_COLOR
 On Error GoTo ErrorManager

    ListBackColor = lstDescripcion.BackColor
    
Exit Property
ErrorManager:
    SubGMuestraError
End Property

Public Property Let ListBackColor(ByVal proColor As OLE_COLOR)
   On Error GoTo ErrorManager

    lstDescripcion.BackColor = proColor
    PropertyChanged "ListBackColor"

Exit Property
ErrorManager:
    SubGMuestraError
       
End Property

Public Property Let ListIndex(Index As Integer)
On Error GoTo ErrorManager

        lstDescripcion.ListIndex = Index
        lstCodigo.ListIndex = Index
        lstAdicional.ListIndex = Index
        PropertyChanged "ListIndex"

Exit Property
ErrorManager:
    SubGMuestraError

    
End Property

Public Property Get ListCount() As Integer

    On Error GoTo ErrorManager
    ListCount = lstDescripcion.ListCount
Exit Property
ErrorManager:
    SubGMuestraError

    
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error GoTo ErrorManager

    PropBag.WriteProperty "BackColor", lstDescripcion.BackColor
    PropBag.WriteProperty "ForeColor", lstDescripcion.ForeColor
    PropBag.WriteProperty "FontInfo", FontInfo
    PropBag.WriteProperty "MostrarBotonesV", bMostrarBotonesV
    PropBag.WriteProperty "MostrarBotonesH", bMostrarBotonesH
    PropBag.WriteProperty "Enabled", bEnabled
    
Exit Sub
ErrorManager:
    SubGMuestraError
    
End Sub

Public Property Get FontInfo() As StdFont
    On Error GoTo ErrorManager

    ' Get the font information
    Set FontInfo = TheFont

Exit Property
ErrorManager:
    SubGMuestraError
    
End Property

Public Property Set FontInfo(NewFont As StdFont)
    On Error GoTo ErrorManager

    ' Set the new font information and then redraw
    Set TheFont = NewFont
    Set lstDescripcion.Font = NewFont
    PropertyChanged "FontInfo"

Exit Property
ErrorManager:
    SubGMuestraError
    
End Property

Public Property Let Texto(proTexto As String)
    On Error GoTo ErrorManager

    lblTexto.Caption = proTexto
    
Exit Property
ErrorManager:
    SubGMuestraError
End Property

