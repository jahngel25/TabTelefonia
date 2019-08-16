VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmRegla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reglas de analisis de digitos"
   ClientHeight    =   10230
   ClientLeft      =   1620
   ClientTop       =   1350
   ClientWidth     =   11400
   Icon            =   "frmRegla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel6 
      Height          =   675
      Left            =   30
      TabIndex        =   11
      Top             =   150
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1191
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Image Image1 
         Height          =   480
         Left            =   90
         Picture         =   "frmRegla.frx":0CCA
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.CheckBox chkVerActivos 
      Caption         =   "&Ver reglas activas solamente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   6
      Top             =   9660
      Value           =   1  'Checked
      Width           =   3795
   End
   Begin VB.Frame fraSeguridad 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   9930
      Width           =   11655
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "Modificar Tramo"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10110
         TabIndex        =   4
         ToolTipText     =   "EliminarTramo"
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Nuevo Tramo"
         Top             =   0
         Width           =   1185
      End
      Begin VB.CommandButton cmdActivar 
         Caption         =   "&Activar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8760
         TabIndex        =   2
         ToolTipText     =   "EliminarTramo"
         Top             =   0
         Width           =   1215
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   705
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9795
      _Version        =   65536
      _ExtentX        =   17277
      _ExtentY        =   1244
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00A7811D&
         Caption         =   $"frmRegla.frx":1994
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   660
         TabIndex        =   10
         Top             =   120
         Width           =   9015
      End
   End
   Begin VB.Frame FrmRegla 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   11415
      Begin MSFlexGridLib.MSFlexGrid grdReglas 
         Height          =   8835
         Left            =   60
         TabIndex        =   7
         Top             =   150
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   15584
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.Label lblCaracteristicas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C09258&
      Caption         =   "Reglas     "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8460
      TabIndex        =   8
      Top             =   90
      Width           =   2955
   End
End
Attribute VB_Name = "frmRegla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
' OBJETIVO: permitir la consulta de las reglas
'****************************************************************
' parametros de entrada: ninguna
' parametros de salida: ninguna
' AUTOR: Hernan Botache
' FECHA: 03/09/2004
'****************************************************************
Option Explicit

'Propiedad de conexion
Public proConexion As ADODB.Connection

'Colección de Reglas
Public proReglas As colRegla
'Bandera de Inicio de Ventana
Dim varBandera As Integer

Function FunFPintaReglas() As Boolean
Dim varContador As Integer
On Error GoTo ErrorManager

    'Adecua la grilla
    grdReglas.Rows = 2
    grdReglas.Cols = 7
    grdReglas.TextMatrix(0, 0) = "ID Regla"
    grdReglas.TextMatrix(0, 1) = "Descripción"
    grdReglas.TextMatrix(0, 2) = "Cant. Digitos"
    grdReglas.TextMatrix(0, 3) = "Repeticiones"
    grdReglas.TextMatrix(0, 4) = "Posic. Digitos"
    grdReglas.TextMatrix(0, 5) = "Consec. Digitos"
    grdReglas.TextMatrix(0, 6) = "RecordStatus"
    grdReglas.FixedRows = 1
    grdReglas.ColWidth(0) = 800
    grdReglas.ColWidth(1) = 8200
    grdReglas.ColWidth(2) = 1100
    grdReglas.ColWidth(3) = 1100
    grdReglas.ColWidth(4) = 1300
    grdReglas.ColWidth(5) = 1300
    grdReglas.ColWidth(6) = 1100
    grdReglas.ColAlignment(0) = 4
    grdReglas.ColAlignment(1) = 4
    grdReglas.ColAlignment(2) = 4
    grdReglas.ColAlignment(3) = 4
    grdReglas.ColAlignment(4) = 4
    grdReglas.ColAlignment(5) = 4
    grdReglas.ColAlignment(6) = 4
    grdReglas.Rows = 1
  


  
    For varContador = 1 To Me.proReglas.Count
        Me.grdReglas.AddItem Me.proReglas.Item(varContador).proReglaId & vbTab & _
                             Trim(Me.proReglas.Item(varContador).proDescripcion) & vbTab & _
                             Me.proReglas.Item(varContador).proCantidadDigitos & vbTab & _
                             Me.proReglas.Item(varContador).proRepeticiones & vbTab & _
                             Me.proReglas.Item(varContador).proPosicionDigitos & vbTab & _
                             Me.proReglas.Item(varContador).proConsecutivoDigitos & vbTab & _
                             Me.proReglas.Item(varContador).proRecordStatus
    
    If Me.proReglas.Item(varContador).proRecordStatus = "0" Then
            SUbFPintaEliminados True, Me.chkVerActivos.Value
        End If
    Next varContador
   Me.grdReglas.Row = 0
   Me.grdReglas.Col = 0
    FunFPintaReglas = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function


Private Sub chkVerActivos_Click()
On Error GoTo ErrorManager

        FunFPintaReglas
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub cmdActivar_Click()
 On Error GoTo ErrManager
    
    If Me.grdReglas.Row = 0 Or Me.grdReglas.RowHeight(Me.grdReglas.RowSel) = 0 Then
        MsgBox "Debe seleccionar la regla a Activar.", vbInformation, App.Title
        Exit Sub
    End If
    
    If Me.proReglas.Item(Me.grdReglas.Row).proRecordStatus = 1 Then
        MsgBox "Este registro ya se encuentra Activado.", vbInformation, App.Title
        Exit Sub
    End If
    
    If MsgBox("Desea Activar la regla: " & Me.proReglas.Item(Me.grdReglas.Row).proDescripcion, vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            Me.proReglas.Item(Me.grdReglas.Row).proRecordStatus = 1
       If Me.proReglas.Item(Me.grdReglas.Row).FunGModificar = True Then
            MsgBox "La regla fue activada exitosamente.", vbInformation, App.Title
        Else
            MsgBox "Error al activar la regla.", vbCritical, App.Title
        End If
    End If
    
    Call FunFPintaReglas
    
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub cmdEliminar_Click()
On Error GoTo ErrorManager

    If Me.grdReglas.Row = 0 Or Me.grdReglas.RowHeight(Me.grdReglas.RowSel) = 0 Then
        MsgBox "Debe seleccionar una regla a eliminar", vbInformation, App.Title
        Exit Sub
    End If
    
    If Me.proReglas.FunGEliminarRegla(Me.grdReglas.Row) = False Then
        MsgBox "No fue posible eliminar la regla", vbInformation + vbOKOnly, App.Title
    End If
    Call FunFPintaReglas
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdModificar_Click()
On Error GoTo ErrorManager

    If (Me.grdReglas.Row = 0 Or Me.grdReglas.RowHeight(Me.grdReglas.RowSel) = 0) Then
        MsgBox "Debe seleccionar una regla a modificar", vbInformation, App.Title
        Exit Sub
    End If
    
    Set frmEdicionReglas.proRegla = Me.proReglas.Item(Me.grdReglas.Row)
        
    frmEdicionReglas.Show vbModal
    Call FunFPintaReglas
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdNuevo_Click()
Dim varRegla As claRegla
On Error GoTo ErrorManager

    Set varRegla = New claRegla
    Set varRegla.proConexion = Me.proConexion
    
    Set frmEdicionReglas.proRegla = varRegla
    
    'Despliegue de ventana de edición
    frmEdicionReglas.Show vbModal
    
    If Len(Trim(varRegla.proReglaId)) <> 0 Then
        'Agrega la recién creada a la colección
        Me.proReglas.Add Me.proConexion, varRegla.proRecordStatus, varRegla.proConsecutivoDigitos, _
        varRegla.proPosicionDigitos, varRegla.proRepeticiones, varRegla.proCantidadDigitos, varRegla.proDescripcion, varRegla.proReglaId
    End If
    
    'Desapliega la colección en la grilla
    Call FunFPintaReglas
    
    Set varRegla = Nothing
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorManager

    If varBandera = 0 Then
        Unload Me
    End If
    varBandera = 2
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
On Error GoTo ErrorManager

    
    varBandera = 1
    
    'Instancia de la colección
    Set Me.proReglas = New colRegla
    Set Me.proReglas.proConexion = Me.proConexion
    
    If Me.proReglas.FunGConsulta = False Then
            MsgBox "No fue posible realizar la consulta de las reglas", vbInformation + vbOKOnly, App.Title
            varBandera = 0
    End If
    
    'Muestra las reglas
    FunFPintaReglas
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorManager

        Set proReglas = Nothing
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Sub SUbFPintaEliminados(parEliminado As Boolean, Optional parOcultar As Variant)
Dim varFila As Integer
Dim varColumna As Integer
On Error GoTo ErrorManager

    Me.grdReglas.Row = Me.grdReglas.Rows - 1
    For varColumna = 0 To Me.grdReglas.Cols - 1
            Me.grdReglas.Col = varColumna
            Me.grdReglas.CellForeColor = &HC0C0C0
    Next varColumna
    
    If parEliminado Then
        If IsMissing(parOcultar) = False Then
            If parOcultar Then
            ' esconde la columna
                Me.grdReglas.RowHeight(Me.grdReglas.Rows - 1) = 0
            End If
        End If
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub


Private Sub grdReglas_DblClick()

On Error GoTo ErrManager
    Call cmdModificar_Click
       Exit Sub
ErrManager:
    SubGMuestraError
End Sub




