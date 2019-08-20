VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEdicionDetalleDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de la información del producto"
   ClientHeight    =   7905
   ClientLeft      =   1875
   ClientTop       =   2130
   ClientWidth     =   11340
   Icon            =   "frmEdicionDetalleDatos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11340
   Begin VB.Frame fraTituloProducto 
      BackColor       =   &H00C09258&
      Caption         =   "  Informaión  del Producto  "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11325
   End
   Begin VB.Frame fraFondoProducto 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   150
      Width           =   11325
      Begin VB.TextBox txtCodigoProducto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2130
         TabIndex        =   3
         Top             =   180
         Width           =   1605
      End
      Begin VB.TextBox txtNombreProducto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3750
         TabIndex        =   2
         Top             =   180
         Width           =   6195
      End
      Begin VB.Label lblProducto 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Left            =   1260
         TabIndex        =   4
         Top             =   210
         Width           =   690
      End
   End
   Begin VB.Frame fraFondoEdicion 
      Height          =   6780
      Left            =   0
      TabIndex        =   5
      Top             =   630
      Width           =   11325
      Begin VB.CheckBox chkValor 
         Caption         =   "Check1"
         Height          =   315
         Index           =   0
         Left            =   2940
         TabIndex        =   15
         Top             =   150
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CommandButton cmdAgregarValores 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   5310
         Picture         =   "frmEdicionDetalleDatos.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Agregar valores"
         Top             =   150
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.ComboBox cboCodigoValor 
         Height          =   315
         Index           =   0
         Left            =   2910
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   150
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboValor 
         Height          =   315
         Index           =   0
         Left            =   2910
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   150
         Visible         =   0   'False
         Width           =   2385
      End
      Begin Threed.SSPanel pnlEtiqueta 
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   150
         Visible         =   0   'False
         Width           =   2865
         _Version        =   65536
         _ExtentX        =   5054
         _ExtentY        =   556
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
      End
      Begin MSComCtl2.DTPicker dtValor 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2910
         TabIndex        =   8
         Top             =   150
         Visible         =   0   'False
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   61407235
         CurrentDate     =   38071
      End
      Begin VB.TextBox txtValor 
         Height          =   315
         Index           =   0
         Left            =   2910
         TabIndex        =   6
         Top             =   150
         Visible         =   0   'False
         Width           =   2385
      End
   End
   Begin VB.Frame fraBotones 
      Height          =   585
      Left            =   0
      TabIndex        =   9
      Top             =   7290
      Width           =   11325
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   180
         Width           =   1425
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   315
         Left            =   9810
         TabIndex        =   10
         Top             =   180
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmEdicionDetalleDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'* Descripcion:
'*
'*
'*
'*
'*
'* Parametros:
'*
'*
'*
'*
'*
'*
'*
'**************************************************************************
'**********************************************************************
' MODIFICADO POR :      CARLOS ALBERTO BARRERA
' DESCRIPCION CAMBIO:   Se pasa como parametro la propiedad del id del cliente
' VERSION: 1.0.100
' FECHA: SEPTIEMBRE 7 /2009
'****************************************************************

Option Explicit

Public proDatosProducto As claDatosProducto
Public proOnyx As EDCVoz.claONYX
Public proParametroProducto As EDCAdminVoz.colParametroProducto
Public proNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
Public proConexion As ADODB.Connection
Public proInsUpd As String
Public proValor As String  'Variable para saber si el activate se ejecuto por la creación  de un nuevo valor

Private varProceso As claProceso
Private varIndiceCboExento 'Se carga con el índice del combo cboValor que indica si está exento de impuesto
Private varConsultando 'Indica que se está cargando la pantalla con la consulta


Public proiClienteId As Long '1.0.100 Propiedad que contiene el id del cliente

Public Enum Categoria
    OT = 1
    Atencion = 2
    Venta = 3
End Enum

Private Sub cboValor_Click(Index As Integer)
    Dim varContador As Integer
    Dim varContadorAux As Integer
    Dim varRegistro As Integer
    Dim varCampo As String
    On Error GoTo ErrManager
    If varConsultando Then Exit Sub
    Me.cboCodigoValor.Item(Index).ListIndex = Me.cboValor.Item(Index).ListIndex
    
    varCampo = Me.cboCodigoValor.Item(Index).Tag
    
    'Deshabilitar el boton de agregar valores
'    For varContador = 1 To Me.cmdAgregarValores.Count
'        Me.cmdAgregarValores.Item(varContador - 1).Enabled = False
'    Next varContador
    
    varRegistro = 0
    For varContador = 1 To Me.proParametroProducto.Count
        If Me.proParametroProducto.Item(varContador).proTipo = "L" Then
        
            If Trim(Me.proParametroProducto.Item(varContador).proCampoPadre) = Trim(varCampo) Then
                Me.proParametroProducto.Item(varContador).proValorIdPadre = Me.cboCodigoValor.Item(Index).Text
              '* 1.0.100 Inicio Se pasa la propiedad del id del cliente
                  If Me.proParametroProducto.Item(varContador).MetConsultarValores(Me.proiClienteId) Then
                '* 1.0.100 Fin
                    
                    Me.cboValor.Item(varRegistro).ListIndex = -1
                    
                    Me.cboCodigoValor.Item(varRegistro).Clear
                    Me.cboValor.Item(varRegistro).Clear
                    
                    For varContadorAux = 1 To Me.proParametroProducto.Item(varContador).proValores.Count
                        Me.cboCodigoValor.Item(varRegistro).AddItem Me.proParametroProducto.Item(varContador).proValores.Item(varContadorAux).proValorID
                        Me.cboValor.Item(varRegistro).AddItem Me.proParametroProducto.Item(varContador).proValores.Item(varContadorAux).proValorDesc
                    Next varContadorAux
                    
                    If Me.proParametroProducto.Item(varContador).proValidarRepetidos = "1" Or Me.proParametroProducto.Item(varContador).proValidarRepetidos = "True" Then
                        Me.cmdAgregarValores.Item(varRegistro).Enabled = True
                    Else
                        Me.cmdAgregarValores.Item(varRegistro).Enabled = False
                    End If
                Else
                    MsgBox "Error al consultar la información de los valores.", vbCritical, App.Title
                    Exit Sub
                End If
               
            End If
            
            If (Trim(Me.proParametroProducto(varContador).proValorIdPadre) <> "" _
                And Trim(Me.proParametroProducto(varContador).proValorIdPadre) <> "0") And _
                (Me.proParametroProducto.Item(varContador).proValidarRepetidos = "1" Or _
                Me.proParametroProducto.Item(varContador).proValidarRepetidos = "True") Then
                
                Me.cmdAgregarValores.Item(varRegistro).Enabled = True
            Else
                Me.cmdAgregarValores.Item(varRegistro).Enabled = False
            End If

            
            varRegistro = varRegistro + 1
        End If
    Next varContador
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdAgregarValores_Click(Index As Integer)
    Dim varContador As Integer
    Dim varContadorValores As Integer
    Dim varCampoPadre1 As String
    Dim varCampoPadre2 As String
    Dim varCampoPadre3 As String
    Dim varValorPadre1 As String
    Dim varValorPadre2 As String
    Dim varPadre As String
    Dim varEntro As Boolean
    Dim varContadorAux As Integer
    Dim varCuentaEntradas As Integer
    Dim varContadorControl As Integer
    On Error GoTo ErrManager
    
    
    For varContador = 1 To Me.proParametroProducto.Count
        If Trim(Me.cmdAgregarValores.Item(Index).Tag) = Trim(Me.proParametroProducto.Item(varContador).proCampo) Then
            Exit For
        End If
    Next
    
    varPadre = Me.proParametroProducto.Item(varContador).proCampoPadre
    varEntro = False
    varCuentaEntradas = 0
    'Buscar si tiene padres o no
    While Trim(varPadre) <> ""
        varCuentaEntradas = varCuentaEntradas + 1
        
        varEntro = True
        
        'Buscar el item padre
        For varContadorAux = 1 To Me.proParametroProducto.Count
            If Trim(Me.proParametroProducto.Item(varContadorAux).proCampo) = Trim(varPadre) Then
                Exit For
            End If
        Next varContadorAux
        
        'buscar el control con el valor padre
        For varContadorControl = 0 To Me.cboCodigoValor.Count - 1
            If Trim(Me.cboCodigoValor.Item(varContadorControl).Tag) = Trim(varPadre) Then
                Exit For
            End If
        Next varContadorControl
        
        If varCuentaEntradas = 1 Then
            varCampoPadre1 = Trim(Me.proParametroProducto.Item(varContadorAux).proCampo)
            varCampoPadre2 = Trim(Me.proParametroProducto.Item(varContador).proCampo)
            varCampoPadre3 = ""
            varValorPadre1 = Trim(Me.cboCodigoValor.Item(varContadorControl).Text)
            varValorPadre2 = ""
        Else
            varCampoPadre3 = Trim(varCampoPadre2)
            varCampoPadre2 = Trim(varCampoPadre1)
            varCampoPadre1 = Trim(Me.proParametroProducto.Item(varContadorAux).proCampo)
            varValorPadre2 = Trim(varValorPadre1)
            varValorPadre1 = Trim(Me.cboCodigoValor.Item(varContadorControl).Text)
        End If
        
        varPadre = Trim(Me.proParametroProducto.Item(varContadorAux).proCampoPadre)
    Wend
        
    If Not varEntro Then
        varCampoPadre1 = Trim(Me.proParametroProducto.Item(varContador).proCampo)
        varCampoPadre2 = ""
        varCampoPadre3 = ""
        varValorPadre1 = ""
        varValorPadre2 = ""
    End If
    
    If Not Me.proParametroProducto.Item(varContador).MetMostrarVentanaEdicion(varCampoPadre1, varCampoPadre2, varCampoPadre3, varValorPadre1, varValorPadre2) Then
        MsgBox "Error al mostrar la ventana de edición de valores."
        Exit Sub
    End If
    
'    Set frmAgregarValor.proConexion = Me.proConexion
'    Set frmAgregarValor.proOnyx = Me.proOnyx
'    Set frmAgregarValor.proNovedadDetalleDatosProducto = Me.proNovedadDetalleDatosProducto
'    Set frmAgregarValor.proParametroProducto = Me.proParametroProducto.Item(varContador)
'    proValor = "S"
'    frmAgregarValor.Show (vbModal)
'
'    Me.cboCodigoValor(Index).Clear
'    Me.cboValor(Index).Clear
'
'    For varContadorValores = 1 To Me.proParametroProducto.Item(varContador).proValores.Count
'        Me.cboValor.Item(Index).AddItem Me.proParametroProducto.Item(varContador).proValores.Item(varContadorValores).proValorDesc
'        Me.cboCodigoValor.Item(Index).AddItem Me.proParametroProducto.Item(varContador).proValores.Item(varContadorValores).proValorID
'    Next varContadorValores
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub CmdCancelar_Click()
    On Error GoTo ErrManager
    
    Unload Me
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdGuardar_Click()
    Dim varContador As Integer
    Dim varContadorAux As Integer
    Dim varPadre As String
    Dim varEncontro As Boolean
    Dim varEtiqueta As String
    Dim varCampoPadre1 As String
    Dim varCampoPadre2 As String
    Dim varCampoPadre3 As String
    Dim varValorPadre1 As String
    Dim varValorPadre2 As String
    Dim varValorPadre3 As String
    Dim varCuentaEntradas As String
    Dim varEntro As Boolean
    Dim varNovedadDetalleDatosProducto As EDCVoz.claNovedadDetalleDatosProducto
    On Error GoTo ErrManager
    
    If Not ValidarCamposObligatorios Then Exit Sub

    
    Set varNovedadDetalleDatosProducto = New EDCVoz.claNovedadDetalleDatosProducto
    Set varNovedadDetalleDatosProducto.proConexion = Me.proConexion
    
    varNovedadDetalleDatosProducto.proRecordStatus = 1
    'Se deben recorrer cada uno de los controles para asignar el valor a las propiedades
    For varContador = 0 To Me.txtValor.Count - 1
        Select Case Trim(Me.txtValor.Item(varContador).Tag)
            Case "vchUser1"
                varNovedadDetalleDatosProducto.proUser1 = Me.txtValor.Item(varContador).Text
            Case "vchUser2"
                varNovedadDetalleDatosProducto.proUser2 = Me.txtValor.Item(varContador).Text
            Case "vchUser3"
                varNovedadDetalleDatosProducto.proUser3 = Me.txtValor.Item(varContador).Text
            Case "vchUser4"
                varNovedadDetalleDatosProducto.proUser4 = Me.txtValor.Item(varContador).Text
            Case "vchUser5"
                varNovedadDetalleDatosProducto.proUser5 = Me.txtValor.Item(varContador).Text
            Case "vchUser6"
                varNovedadDetalleDatosProducto.proUser6 = Me.txtValor.Item(varContador).Text
            Case "vchUser7"
                varNovedadDetalleDatosProducto.proUser7 = Me.txtValor.Item(varContador).Text
            Case "vchUser8"
                varNovedadDetalleDatosProducto.proUser8 = Me.txtValor.Item(varContador).Text
            Case "vchUser9"
                varNovedadDetalleDatosProducto.proUser9 = Me.txtValor.Item(varContador).Text
            Case "vchUser10"
                varNovedadDetalleDatosProducto.proUser10 = Me.txtValor.Item(varContador).Text
            Case "vchUser11"
                varNovedadDetalleDatosProducto.proUser11 = Me.txtValor.Item(varContador).Text
            Case "vchUser12"
                varNovedadDetalleDatosProducto.proUser12 = Me.txtValor.Item(varContador).Text
            Case "vchUser13"
                varNovedadDetalleDatosProducto.proUser13 = Me.txtValor.Item(varContador).Text
            Case "vchUser14"
                varNovedadDetalleDatosProducto.proUser14 = Me.txtValor.Item(varContador).Text
            Case "vchUser15"
                varNovedadDetalleDatosProducto.proUser15 = Me.txtValor.Item(varContador).Text
            Case "vchUser16"
                varNovedadDetalleDatosProducto.proUser16 = Me.txtValor.Item(varContador).Text
            Case "vchUser17"
                varNovedadDetalleDatosProducto.proUser17 = Me.txtValor.Item(varContador).Text
            Case "vchUser18"
                varNovedadDetalleDatosProducto.proUser18 = Me.txtValor.Item(varContador).Text
            Case "vchUser19"
                varNovedadDetalleDatosProducto.proUser19 = Me.txtValor.Item(varContador).Text
            Case "vchUser20"
                varNovedadDetalleDatosProducto.proUser20 = Me.txtValor.Item(varContador).Text
            Case "vchUser21"
                varNovedadDetalleDatosProducto.proUser21 = Me.txtValor.Item(varContador).Text
            Case "vchUser22"
                varNovedadDetalleDatosProducto.proUser22 = Me.txtValor.Item(varContador).Text
            Case "vchUser23"
                varNovedadDetalleDatosProducto.proUser23 = Me.txtValor.Item(varContador).Text
            Case "vchUser24"
                varNovedadDetalleDatosProducto.proUser24 = Me.txtValor.Item(varContador).Text
            Case "vchUser25"
                varNovedadDetalleDatosProducto.proUser25 = Me.txtValor.Item(varContador).Text
            Case "vchUser26"
                varNovedadDetalleDatosProducto.proUser26 = Me.txtValor.Item(varContador).Text
            Case "vchUser27"
                varNovedadDetalleDatosProducto.proUser27 = Me.txtValor.Item(varContador).Text
            Case "vchUser28"
                varNovedadDetalleDatosProducto.proUser28 = Me.txtValor.Item(varContador).Text
            Case "vchUser29"
                varNovedadDetalleDatosProducto.proUser29 = Me.txtValor.Item(varContador).Text
            Case "vchUser30"
                varNovedadDetalleDatosProducto.proUser30 = Me.txtValor.Item(varContador).Text
            Case "vchUser31"
                varNovedadDetalleDatosProducto.proUser31 = Me.txtValor.Item(varContador).Text
            Case "vchUser32"
                varNovedadDetalleDatosProducto.proUser32 = Me.txtValor.Item(varContador).Text
            Case "vchUser33"
                varNovedadDetalleDatosProducto.proUser33 = Me.txtValor.Item(varContador).Text
            Case "vchUser34"
                varNovedadDetalleDatosProducto.proUser34 = Me.txtValor.Item(varContador).Text
            Case "vchUser35"
                varNovedadDetalleDatosProducto.proUser35 = Me.txtValor.Item(varContador).Text
            Case "vchUser36"
                varNovedadDetalleDatosProducto.proUser36 = Me.txtValor.Item(varContador).Text
            Case "vchUser37"
                varNovedadDetalleDatosProducto.proUser37 = Me.txtValor.Item(varContador).Text
            Case "vchUser38"
                varNovedadDetalleDatosProducto.proUser38 = Me.txtValor.Item(varContador).Text
            Case "vchUser39"
                varNovedadDetalleDatosProducto.proUser39 = Me.txtValor.Item(varContador).Text
            Case "vchUser40"
                varNovedadDetalleDatosProducto.proUser40 = Me.txtValor.Item(varContador).Text
        End Select
    Next varContador
    
    'Se deben recorrer cada uno de los controles para asignar el valor a las propiedades
    For varContador = 0 To Me.dtValor.Count - 1
        Select Case Trim(Me.dtValor.Item(varContador).Tag)
            Case "vchUser1"
                varNovedadDetalleDatosProducto.proUser1 = Me.dtValor.Item(varContador).Value
            Case "vchUser2"
                varNovedadDetalleDatosProducto.proUser2 = Me.dtValor.Item(varContador).Value
            Case "vchUser3"
                varNovedadDetalleDatosProducto.proUser3 = Me.dtValor.Item(varContador).Value
            Case "vchUser4"
                varNovedadDetalleDatosProducto.proUser4 = Me.dtValor.Item(varContador).Value
            Case "vchUser5"
                varNovedadDetalleDatosProducto.proUser5 = Me.dtValor.Item(varContador).Value
            Case "vchUser6"
                varNovedadDetalleDatosProducto.proUser6 = Me.dtValor.Item(varContador).Value
            Case "vchUser7"
                varNovedadDetalleDatosProducto.proUser7 = Me.dtValor.Item(varContador).Value
            Case "vchUser8"
                varNovedadDetalleDatosProducto.proUser8 = Me.dtValor.Item(varContador).Value
            Case "vchUser9"
                varNovedadDetalleDatosProducto.proUser9 = Me.dtValor.Item(varContador).Value
            Case "vchUser10"
                varNovedadDetalleDatosProducto.proUser10 = Me.dtValor.Item(varContador).Value
            Case "vchUser11"
                varNovedadDetalleDatosProducto.proUser11 = Me.dtValor.Item(varContador).Value
            Case "vchUser12"
                varNovedadDetalleDatosProducto.proUser12 = Me.dtValor.Item(varContador).Value
            Case "vchUser13"
                varNovedadDetalleDatosProducto.proUser13 = Me.dtValor.Item(varContador).Value
            Case "vchUser14"
                varNovedadDetalleDatosProducto.proUser14 = Me.dtValor.Item(varContador).Value
            Case "vchUser15"
                varNovedadDetalleDatosProducto.proUser15 = Me.dtValor.Item(varContador).Value
            Case "vchUser16"
                varNovedadDetalleDatosProducto.proUser16 = Me.dtValor.Item(varContador).Value
            Case "vchUser17"
                varNovedadDetalleDatosProducto.proUser17 = Me.dtValor.Item(varContador).Value
            Case "vchUser18"
                varNovedadDetalleDatosProducto.proUser18 = Me.dtValor.Item(varContador).Value
            Case "vchUser19"
                varNovedadDetalleDatosProducto.proUser19 = Me.dtValor.Item(varContador).Value
            Case "vchUser20"
                varNovedadDetalleDatosProducto.proUser20 = Me.dtValor.Item(varContador).Value
            Case "vchUser21"
                varNovedadDetalleDatosProducto.proUser21 = Me.dtValor.Item(varContador).Value
            Case "vchUser22"
                varNovedadDetalleDatosProducto.proUser22 = Me.dtValor.Item(varContador).Value
            Case "vchUser23"
                varNovedadDetalleDatosProducto.proUser23 = Me.dtValor.Item(varContador).Value
            Case "vchUser24"
                varNovedadDetalleDatosProducto.proUser24 = Me.dtValor.Item(varContador).Value
            Case "vchUser25"
                varNovedadDetalleDatosProducto.proUser25 = Me.dtValor.Item(varContador).Value
            Case "vchUser26"
                varNovedadDetalleDatosProducto.proUser26 = Me.dtValor.Item(varContador).Value
            Case "vchUser27"
                varNovedadDetalleDatosProducto.proUser27 = Me.dtValor.Item(varContador).Value
            Case "vchUser28"
                varNovedadDetalleDatosProducto.proUser28 = Me.dtValor.Item(varContador).Value
            Case "vchUser29"
                varNovedadDetalleDatosProducto.proUser29 = Me.dtValor.Item(varContador).Value
            Case "vchUser30"
                varNovedadDetalleDatosProducto.proUser30 = Me.dtValor.Item(varContador).Value
            Case "vchUser31"
                varNovedadDetalleDatosProducto.proUser31 = Me.dtValor.Item(varContador).Value
            Case "vchUser32"
                varNovedadDetalleDatosProducto.proUser32 = Me.dtValor.Item(varContador).Value
            Case "vchUser33"
                varNovedadDetalleDatosProducto.proUser33 = Me.dtValor.Item(varContador).Value
            Case "vchUser34"
                varNovedadDetalleDatosProducto.proUser34 = Me.dtValor.Item(varContador).Value
            Case "vchUser35"
                varNovedadDetalleDatosProducto.proUser35 = Me.dtValor.Item(varContador).Value
            Case "vchUser36"
                varNovedadDetalleDatosProducto.proUser36 = Me.dtValor.Item(varContador).Value
            Case "vchUser37"
                varNovedadDetalleDatosProducto.proUser37 = Me.dtValor.Item(varContador).Value
            Case "vchUser38"
                varNovedadDetalleDatosProducto.proUser38 = Me.dtValor.Item(varContador).Value
            Case "vchUser39"
                varNovedadDetalleDatosProducto.proUser39 = Me.dtValor.Item(varContador).Value
            Case "vchUser40"
                varNovedadDetalleDatosProducto.proUser40 = Me.dtValor.Item(varContador).Value
        End Select
    Next varContador
    
    For varContador = 0 To Me.cboValor.Count - 1
        Select Case Trim(Me.cboValor.Item(varContador).Tag)
            Case "vchUser1"
                varNovedadDetalleDatosProducto.proUser1 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser2"
                varNovedadDetalleDatosProducto.proUser2 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser3"
                varNovedadDetalleDatosProducto.proUser3 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser4"
                varNovedadDetalleDatosProducto.proUser4 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser5"
                varNovedadDetalleDatosProducto.proUser5 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser6"
                varNovedadDetalleDatosProducto.proUser6 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser7"
                varNovedadDetalleDatosProducto.proUser7 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser8"
                varNovedadDetalleDatosProducto.proUser8 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser9"
                varNovedadDetalleDatosProducto.proUser9 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser10"
                varNovedadDetalleDatosProducto.proUser10 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser11"
                varNovedadDetalleDatosProducto.proUser11 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser12"
                varNovedadDetalleDatosProducto.proUser12 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser13"
                varNovedadDetalleDatosProducto.proUser13 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser14"
                varNovedadDetalleDatosProducto.proUser14 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser15"
                varNovedadDetalleDatosProducto.proUser15 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser16"
                varNovedadDetalleDatosProducto.proUser16 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser17"
                varNovedadDetalleDatosProducto.proUser17 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser18"
                varNovedadDetalleDatosProducto.proUser18 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser19"
                varNovedadDetalleDatosProducto.proUser19 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser20"
                varNovedadDetalleDatosProducto.proUser20 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser21"
                varNovedadDetalleDatosProducto.proUser21 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser22"
                varNovedadDetalleDatosProducto.proUser22 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser23"
                varNovedadDetalleDatosProducto.proUser23 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser24"
                varNovedadDetalleDatosProducto.proUser24 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser25"
                varNovedadDetalleDatosProducto.proUser25 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser26"
                varNovedadDetalleDatosProducto.proUser26 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser27"
                varNovedadDetalleDatosProducto.proUser27 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser28"
                varNovedadDetalleDatosProducto.proUser28 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser29"
                varNovedadDetalleDatosProducto.proUser29 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser30"
                varNovedadDetalleDatosProducto.proUser30 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser31"
                varNovedadDetalleDatosProducto.proUser31 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser32"
                varNovedadDetalleDatosProducto.proUser32 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser33"
                varNovedadDetalleDatosProducto.proUser33 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser34"
                varNovedadDetalleDatosProducto.proUser34 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser35"
                varNovedadDetalleDatosProducto.proUser35 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser36"
                varNovedadDetalleDatosProducto.proUser36 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser37"
                varNovedadDetalleDatosProducto.proUser37 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser38"
                varNovedadDetalleDatosProducto.proUser38 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser39"
                varNovedadDetalleDatosProducto.proUser39 = Me.cboCodigoValor.Item(varContador).Text
            Case "vchUser40"
                varNovedadDetalleDatosProducto.proUser40 = Me.cboCodigoValor.Item(varContador).Text
        End Select
    Next varContador
    
    'Agregando campos tipo check box
    For varContador = 0 To Me.chkValor.Count - 1
        Select Case Trim(Me.chkValor.Item(varContador).Tag)
            Case "vchUser1"
                varNovedadDetalleDatosProducto.proUser1 = Me.chkValor.Item(varContador).Value
            Case "vchUser2"
                varNovedadDetalleDatosProducto.proUser2 = Me.chkValor.Item(varContador).Value
            Case "vchUser3"
                varNovedadDetalleDatosProducto.proUser3 = Me.chkValor.Item(varContador).Value
            Case "vchUser4"
                varNovedadDetalleDatosProducto.proUser4 = Me.chkValor.Item(varContador).Value
            Case "vchUser5"
                varNovedadDetalleDatosProducto.proUser5 = Me.chkValor.Item(varContador).Value
            Case "vchUser6"
                varNovedadDetalleDatosProducto.proUser6 = Me.chkValor.Item(varContador).Value
            Case "vchUser7"
                varNovedadDetalleDatosProducto.proUser7 = Me.chkValor.Item(varContador).Value
            Case "vchUser8"
                varNovedadDetalleDatosProducto.proUser8 = Me.chkValor.Item(varContador).Value
            Case "vchUser9"
                varNovedadDetalleDatosProducto.proUser9 = Me.chkValor.Item(varContador).Value
            Case "vchUser10"
                varNovedadDetalleDatosProducto.proUser10 = Me.chkValor.Item(varContador).Value
            Case "vchUser11"
                varNovedadDetalleDatosProducto.proUser11 = Me.chkValor.Item(varContador).Value
            Case "vchUser12"
                varNovedadDetalleDatosProducto.proUser12 = Me.chkValor.Item(varContador).Value
            Case "vchUser13"
                varNovedadDetalleDatosProducto.proUser13 = Me.chkValor.Item(varContador).Value
            Case "vchUser14"
                varNovedadDetalleDatosProducto.proUser14 = Me.chkValor.Item(varContador).Value
            Case "vchUser15"
                varNovedadDetalleDatosProducto.proUser15 = Me.chkValor.Item(varContador).Value
            Case "vchUser16"
                varNovedadDetalleDatosProducto.proUser16 = Me.chkValor.Item(varContador).Value
            Case "vchUser17"
                varNovedadDetalleDatosProducto.proUser17 = Me.chkValor.Item(varContador).Value
            Case "vchUser18"
                varNovedadDetalleDatosProducto.proUser18 = Me.chkValor.Item(varContador).Value
            Case "vchUser19"
                varNovedadDetalleDatosProducto.proUser19 = Me.chkValor.Item(varContador).Value
            Case "vchUser20"
                varNovedadDetalleDatosProducto.proUser20 = Me.chkValor.Item(varContador).Value
            Case "vchUser21"
                varNovedadDetalleDatosProducto.proUser21 = Me.chkValor.Item(varContador).Value
            Case "vchUser22"
                varNovedadDetalleDatosProducto.proUser22 = Me.chkValor.Item(varContador).Value
            Case "vchUser23"
                varNovedadDetalleDatosProducto.proUser23 = Me.chkValor.Item(varContador).Value
            Case "vchUser24"
                varNovedadDetalleDatosProducto.proUser24 = Me.chkValor.Item(varContador).Value
            Case "vchUser25"
                varNovedadDetalleDatosProducto.proUser25 = Me.chkValor.Item(varContador).Value
            Case "vchUser26"
                varNovedadDetalleDatosProducto.proUser26 = Me.chkValor.Item(varContador).Value
            Case "vchUser27"
                varNovedadDetalleDatosProducto.proUser27 = Me.chkValor.Item(varContador).Value
            Case "vchUser28"
                varNovedadDetalleDatosProducto.proUser28 = Me.chkValor.Item(varContador).Value
            Case "vchUser29"
                varNovedadDetalleDatosProducto.proUser29 = Me.chkValor.Item(varContador).Value
            Case "vchUser30"
                varNovedadDetalleDatosProducto.proUser30 = Me.chkValor.Item(varContador).Value
            Case "vchUser31"
                varNovedadDetalleDatosProducto.proUser31 = Me.chkValor.Item(varContador).Value
            Case "vchUser32"
                varNovedadDetalleDatosProducto.proUser32 = Me.chkValor.Item(varContador).Value
            Case "vchUser33"
                varNovedadDetalleDatosProducto.proUser33 = Me.chkValor.Item(varContador).Value
            Case "vchUser34"
                varNovedadDetalleDatosProducto.proUser34 = Me.chkValor.Item(varContador).Value
            Case "vchUser35"
                varNovedadDetalleDatosProducto.proUser35 = Me.chkValor.Item(varContador).Value
            Case "vchUser36"
                varNovedadDetalleDatosProducto.proUser36 = Me.chkValor.Item(varContador).Value
            Case "vchUser37"
                varNovedadDetalleDatosProducto.proUser37 = Me.chkValor.Item(varContador).Value
            Case "vchUser38"
                varNovedadDetalleDatosProducto.proUser38 = Me.chkValor.Item(varContador).Value
            Case "vchUser39"
                varNovedadDetalleDatosProducto.proUser39 = Me.chkValor.Item(varContador).Value
            Case "vchUser40"
                varNovedadDetalleDatosProducto.proUser40 = Me.chkValor.Item(varContador).Value
        End Select
    Next varContador
    
    Set varNovedadDetalleDatosProducto.proParametrosxProducto = Me.proParametroProducto
    
    'Determinar la norma
    If Me.proDatosProducto.proTipoTelefonia <> "107441" Then
        Dim varExento As Boolean: varExento = False
        If varIndiceCboExento <> 1 Then varExento = (cboValor(varIndiceCboExento).Text = "Si")
        If varExento Then 'Si está exento, aplica la norma parametrizada
            Dim varClaParametro As New claParametro
            Set varClaParametro.proConexion = Me.proConexion
            varClaParametro.proAcronimo = "CodigoExento"
            varClaParametro.FunGConsultar
            varNovedadDetalleDatosProducto.proUser17 = ""
            If InStr(1, varClaParametro.proValor, "-") = 0 Then
                varNovedadDetalleDatosProducto.proUser18 = varClaParametro.proValor
            Else
                varNovedadDetalleDatosProducto.proUser18 = Mid(varClaParametro.proValor, 1, InStr(1, varClaParametro.proValor, "-") - 1)
            End If
        Else
            Dim varColNorma As New ColNorma
            Set varColNorma.proConexion = Me.proConexion
            varColNorma.FunGConsultaporParametros Val(varNovedadDetalleDatosProducto.proUser8), proDatosProducto.proUsoServicioId, Val(varNovedadDetalleDatosProducto.proUser1), Val(proDatosProducto.proiEstratoid)
            If varColNorma.Count = 0 Then
                varNovedadDetalleDatosProducto.proUser17 = ""
                varNovedadDetalleDatosProducto.proUser18 = ""
            Else
                varNovedadDetalleDatosProducto.proUser17 = varColNorma.Item(1).proNormaId
                varNovedadDetalleDatosProducto.proUser18 = varColNorma.Item(1).proCodigoNorma
            End If
        End If
    End If
    'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
    If Not Me.proDatosProducto.MetGuardar Then
        MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
        Exit Sub
    End If
    
    'Inserta o actualiza la información de los incidentes
    If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
        MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
        Exit Sub
    End If
    
    'Buscar los campos que se les debe validar la existencia
    For varContador = 1 To Me.proParametroProducto.Count
        If Me.proParametroProducto.Item(varContador).proValidarRepetidos = 1 Then
            
            'Buscar los padres y los valores respectivos
            varPadre = Trim(Me.proParametroProducto.Item(varContador).proCampoPadre)
            varEntro = False
            varCuentaEntradas = 0
            While Trim(varPadre) <> ""
                varEntro = True
                varCuentaEntradas = varCuentaEntradas + 1
                
                'Buscar el padre
                For varContadorAux = 1 To Me.proParametroProducto.Count
                    If Trim(Me.proParametroProducto.Item(varContadorAux).proCampo) = Trim(varPadre) Then
                        Exit For
                    End If
                Next varContadorAux
                
                If varCuentaEntradas = 1 Then
                    varCampoPadre1 = Trim(Me.proParametroProducto.Item(varContadorAux).proCampo)
                    varCampoPadre2 = Trim(Me.proParametroProducto.Item(varContador).proCampo)
                    varCampoPadre3 = ""
                    
                    'Asignar el valor del padre 1
                    Select Case Trim(varCampoPadre1)
                        Case "vchUser1"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser1
                        Case "vchUser2"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser2
                        Case "vchUser3"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser3
                        Case "vchUser4"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser4
                        Case "vchUser5"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser5
                        Case "vchUser6"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser6
                        Case "vchUser7"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser7
                        Case "vchUser8"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser8
                        Case "vchUser9"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser9
                        Case "vchUser10"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser10
                        Case "vchUser11"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser11
                        Case "vchUser12"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser12
                        Case "vchUser13"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser13
                        Case "vchUser14"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser14
                        Case "vchUser15"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser15
                        Case "vchUser16"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser16
                        Case "vchUser7"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser17
                        Case "vchUser18"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser18
                        Case "vchUser19"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser19
                        Case "vchUser20"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser20
                        Case "vchUser21"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser21
                        Case "vchUser22"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser22
                        Case "vchUser23"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser23
                        Case "vchUser24"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser24
                        Case "vchUser25"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser25
                        Case "vchUser26"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser26
                        Case "vchUser27"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser27
                        Case "vchUser28"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser28
                        Case "vchUser29"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser29
                        Case "vchUser30"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser30
                        Case "vchUser31"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser31
                        Case "vchUser32"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser32
                        Case "vchUser33"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser33
                        Case "vchUser34"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser34
                        Case "vchUser35"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser35
                        Case "vchUser36"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser36
                        Case "vchUser37"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser37
                        Case "vchUser38"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser38
                        Case "vchUser39"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser39
                        Case "vchUser40"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser40
                    End Select
                
                    'Asignar el valor del padre 2
                    Select Case Trim(varCampoPadre2)
                        Case "vchUser1"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser1
                        Case "vchUser2"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser2
                        Case "vchUser3"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser3
                        Case "vchUser4"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser4
                        Case "vchUser5"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser5
                        Case "vchUser6"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser6
                        Case "vchUser7"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser7
                        Case "vchUser8"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser8
                        Case "vchUser9"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser9
                        Case "vchUser10"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser10
                        Case "vchUser11"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser11
                        Case "vchUser12"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser12
                        Case "vchUser13"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser13
                        Case "vchUser14"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser14
                        Case "vchUser15"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser15
                        Case "vchUser16"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser16
                        Case "vchUser7"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser17
                        Case "vchUser18"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser18
                        Case "vchUser19"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser19
                        Case "vchUser20"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser20
                        Case "vchUser21"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser21
                        Case "vchUser22"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser22
                        Case "vchUser23"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser23
                        Case "vchUser24"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser24
                        Case "vchUser25"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser25
                        Case "vchUser26"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser26
                        Case "vchUser27"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser27
                        Case "vchUser28"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser28
                        Case "vchUser29"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser29
                        Case "vchUser30"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser30
                        Case "vchUser31"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser31
                        Case "vchUser32"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser32
                        Case "vchUser33"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser33
                        Case "vchUser34"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser34
                        Case "vchUser35"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser35
                        Case "vchUser36"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser36
                        Case "vchUser37"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser37
                        Case "vchUser38"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser38
                        Case "vchUser39"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser39
                        Case "vchUser40"
                            varValorPadre2 = varNovedadDetalleDatosProducto.proUser40
                    End Select
                    
                    varValorPadre3 = ""
                Else
                    varCampoPadre3 = Trim(varCampoPadre2)
                    varCampoPadre2 = Trim(varCampoPadre1)
                    varCampoPadre1 = Trim(Me.proParametroProducto.Item(varContadorAux).proCampo)
                    
                    varValorPadre3 = Trim(varValorPadre2)
                    varValorPadre2 = Trim(varValorPadre1)
                    
                    'Asignar el valor del padre 1
                    Select Case Trim(varCampoPadre1)
                        Case "vchUser1"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser1
                        Case "vchUser2"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser2
                        Case "vchUser3"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser3
                        Case "vchUser4"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser4
                        Case "vchUser5"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser5
                        Case "vchUser6"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser6
                        Case "vchUser7"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser7
                        Case "vchUser8"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser8
                        Case "vchUser9"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser9
                        Case "vchUser10"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser10
                        Case "vchUser11"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser11
                        Case "vchUser12"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser12
                        Case "vchUser13"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser13
                        Case "vchUser14"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser14
                        Case "vchUser15"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser15
                        Case "vchUser16"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser16
                        Case "vchUser7"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser17
                        Case "vchUser18"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser18
                        Case "vchUser19"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser19
                        Case "vchUser20"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser20
                        Case "vchUser21"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser21
                        Case "vchUser22"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser22
                        Case "vchUser23"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser23
                        Case "vchUser24"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser24
                        Case "vchUser25"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser25
                        Case "vchUser26"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser26
                        Case "vchUser27"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser27
                        Case "vchUser28"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser28
                        Case "vchUser29"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser29
                        Case "vchUser30"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser30
                        Case "vchUser31"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser31
                        Case "vchUser32"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser32
                        Case "vchUser33"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser33
                        Case "vchUser34"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser34
                        Case "vchUser35"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser35
                        Case "vchUser36"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser36
                        Case "vchUser37"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser37
                        Case "vchUser38"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser38
                        Case "vchUser39"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser39
                        Case "vchUser40"
                            varValorPadre1 = varNovedadDetalleDatosProducto.proUser40
                    End Select
                End If
                varPadre = Trim(Me.proParametroProducto.Item(varContadorAux).proCampoPadre)
            Wend
            
            If Not varEntro Then
                varCampoPadre1 = Trim(Me.proParametroProducto.Item(varContador).proCampo)
                varCampoPadre2 = ""
                varCampoPadre3 = ""
                Select Case Trim(varCampoPadre1)
                    Case "vchUser1"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser1
                    Case "vchUser2"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser2
                    Case "vchUser3"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser3
                    Case "vchUser4"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser4
                    Case "vchUser5"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser5
                    Case "vchUser6"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser6
                    Case "vchUser7"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser7
                    Case "vchUser8"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser8
                    Case "vchUser9"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser9
                    Case "vchUser10"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser10
                    Case "vchUser11"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser11
                    Case "vchUser12"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser12
                    Case "vchUser13"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser13
                    Case "vchUser14"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser14
                    Case "vchUser15"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser15
                    Case "vchUser16"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser16
                    Case "vchUser7"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser17
                    Case "vchUser18"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser18
                    Case "vchUser19"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser19
                    Case "vchUser20"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser20
                    Case "vchUser21"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser21
                    Case "vchUser22"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser22
                    Case "vchUser23"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser23
                    Case "vchUser24"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser24
                    Case "vchUser25"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser25
                    Case "vchUser26"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser26
                    Case "vchUser27"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser27
                    Case "vchUser28"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser28
                    Case "vchUser29"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser29
                    Case "vchUser30"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser30
                    Case "vchUser31"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser31
                    Case "vchUser32"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser32
                    Case "vchUser33"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser33
                    Case "vchUser34"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser34
                    Case "vchUser35"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser35
                    Case "vchUser36"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser36
                    Case "vchUser37"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser37
                    Case "vchUser38"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser38
                    Case "vchUser39"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser39
                    Case "vchUser40"
                        varValorPadre1 = varNovedadDetalleDatosProducto.proUser40
                End Select
            End If
            
            If Trim(varValorPadre1) <> "" And varValorPadre1 <> "0" Then
            
                If Not Me.proParametroProducto.Item(varContador).MetValidarInformacionCampo(varCampoPadre1, varCampoPadre2, varCampoPadre3, varValorPadre1, varValorPadre2, varValorPadre3, Me.proDatosProducto.proDatosProductoId, varEtiqueta) Then
                    MsgBox "El campo [" & varEtiqueta & "] tiene un valor que ya fue usado en otro servicio.", vbInformation, App.Title
                    Exit Sub
                End If
            End If
        End If
    Next varContador
    
    
    'Guardar la información del detalle
    varNovedadDetalleDatosProducto.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
    varNovedadDetalleDatosProducto.proDetalleDatosProductoId = Me.proNovedadDetalleDatosProducto.proDetalleDatosProductoId

    If Trim(varNovedadDetalleDatosProducto.proDetalleDatosProductoId) = "" Then
        varNovedadDetalleDatosProducto.proDetalleDatosProductoId = 0
    End If
    varNovedadDetalleDatosProducto.proIncidentId = Me.proDatosProducto.proIncidentId
    varNovedadDetalleDatosProducto.proStatusId = "A"
    varNovedadDetalleDatosProducto.proTipoNovedadId = Me.proNovedadDetalleDatosProducto.proTipoNovedadId
    varNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId = Me.proNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId
            
    If varNovedadDetalleDatosProducto.MetGuardar Then
        
        varEncontro = False
        
        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
            If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proNovedadDetalleDatosProductoId = _
                varNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId Then
                varEncontro = True
                Exit For
            End If
        Next varContador
        
        If varEncontro = False Then
            If Me.proDatosProducto.MetAgregarNovedadDetalle(varNovedadDetalleDatosProducto) Then
                MsgBox "La información se almacenó exitosamente.", vbInformation, App.Title
                Unload Me
            Else
                MsgBox "Error al almacenar la información de los detalles.", vbCritical, App.Title
            End If
        Else
            If Me.proDatosProducto.MetActualizarNovedadDetalle(varNovedadDetalleDatosProducto) Then
                MsgBox "La información se almacenó exitosamente.", vbInformation, App.Title
                Unload Me
            Else
                MsgBox "Error al almacenar la información de los detalles.", vbCritical, App.Title
            End If
        End If
    Else
        MsgBox "Error al almacenar la información de los detalles.", vbCritical, App.Title
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub Form_Activate()
    On Error GoTo ErrManager
        
    'Si se crea un valor no debe volver a repintar la pantalla
    If proValor = "S" Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    varConsultando = True
    Call SubFInicializarPantalla
    varConsultando = False
    
    If Me.proInsUpd = "U" Then
        Call SubFCargarValores
    Else
        Me.proNovedadDetalleDatosProducto.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
        Me.proNovedadDetalleDatosProducto.proStatusId = 0
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub SubFInicializarPantalla()
    Dim varContador As Integer
    Dim varCantidad As Integer
    Dim varIndice As Integer
    Dim varTop As Integer
    Dim varLeftPnl As Integer
    Dim varLeftVal As Integer
    Dim varLeftBtn As Integer
    Dim varContadorValores As Integer
    Dim varCantidadBotonesHabilitados As Integer
    On Error GoTo ErrManager
    
    Me.txtCodigoProducto.Text = Me.proDatosProducto.proProductNumber
    Me.txtNombreProducto.Text = Me.proDatosProducto.proProductName
    varCantidadBotonesHabilitados = 0
    varIndiceCboExento = -1
    
    If Me.proParametroProducto.Count = 0 Then
        MsgBox "El producto no tiene parametros configurados.", vbInformation, App.Title
        Exit Sub
    End If
    
    'Limpiar los controles
    For varContador = 1 To Me.txtValor.Count - 1
        Unload Me.txtValor.Item(varContador)
    Next varContador
    
    For varContador = 1 To Me.cboCodigoValor.Count - 1
        Unload Me.cboCodigoValor.Item(varContador)
        Unload Me.cboValor.Item(varContador)
    Next varContador
    
    For varContador = 1 To Me.dtValor.Count - 1
        Unload Me.dtValor.Item(varContador)
    Next varContador
    
    For varContador = 1 To Me.cmdAgregarValores.Count - 1
        Unload Me.cmdAgregarValores.Item(varContador)
    Next varContador
    
    For varContador = 1 To Me.chkValor.Count - 1
        Unload Me.chkValor.Item(varContador)
    Next varContador
    
    Me.txtValor.Item(0).Visible = False
    Me.dtValor.Item(0).Visible = False
    Me.cboCodigoValor.Item(0).Visible = False
    Me.cboValor.Item(0).Visible = False
    Me.cmdAgregarValores.Item(0).Visible = False
    Me.chkValor.Item(0).Visible = False
    
    'Consultar los datos del incidente
    Set varProceso = New claProceso
    Set varProceso.proConexion = Me.proConexion
    
    varProceso.proIncidentId = Me.proDatosProducto.proIncidentId
    
    If Not varProceso.MetConsultaDatosIncidente Then
        MsgBox "Error al buscar la información del incidente.", vbCritical, App.Title
        Exit Sub
    End If
    
    'Pintar la forma
    For varContador = 0 To Me.proParametroProducto.Count - 1
        If varContador < 20 Then
            varTop = varContador
            varLeftPnl = 30
            varLeftVal = 2910
            varLeftBtn = 5310
        Else
            varTop = varContador - 20
            varLeftPnl = 5670
            varLeftVal = 8550
            varLeftBtn = 10950
        End If
        
        If varContador = 0 Then
            Me.pnlEtiqueta.Item(varContador).Visible = True
            Me.pnlEtiqueta.Item(varContador).Caption = Me.proParametroProducto.Item(varContador + 1).proEtiqueta
            Me.pnlEtiqueta.Item(varContador).Left = varLeftPnl
            Me.pnlEtiqueta.Item(varContador).Top = 150
            
            Select Case Me.proParametroProducto.Item(varContador + 1).proTipo
                Case "T"
                    Me.txtValor.Item(varContador).Visible = True
                    Me.txtValor.Item(varContador).Top = 150
                    Me.txtValor.Item(varContador).Left = varLeftVal
                    Me.txtValor.Item(varContador).Text = ""
                    Me.txtValor.Item(varContador).MaxLength = Me.proParametroProducto.Item(varContador + 1).proTamaño
                    Me.txtValor.Item(varContador).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    If FunFCamposObligatorios(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Or FunFCamposEditables(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Then
                        Me.txtValor.Item(varContador).Enabled = True
                    Else
                        Me.txtValor.Item(varContador).Enabled = True
                    End If
                    Me.txtValor.Item(varContador).TabIndex = varContador
                    If Me.txtValor.Item(varContador).Enabled = True Then
                        Me.txtValor.Item(varContador).SetFocus
                    End If
                Case "F"
                    Me.dtValor.Item(varContador).Visible = True
                    Me.dtValor.Item(varContador).Top = 150
                    Me.dtValor.Item(varContador).Left = varLeftVal
                    Me.dtValor.Item(varContador).Value = Now
                    Me.dtValor.Item(varContador).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    If FunFCamposObligatorios(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Or FunFCamposEditables(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Then
                        Me.dtValor.Item(varContador).Enabled = True
                    Else
                        Me.dtValor.Item(varContador).Enabled = False
                    End If
                    Me.dtValor.Item(varContador).TabIndex = varContador
                    If Me.dtValor.Item(varContador).Enabled = True Then
                        Me.dtValor.Item(varContador).SetFocus
                    End If
                Case "L"
                    Me.cboValor.Item(varContador).Visible = True
                    Me.cboValor.Item(varContador).Top = 150
                    Me.cboValor.Item(varContador).Left = varLeftVal
                    Me.cboValor.Item(varContador).ListIndex = -1
                    Me.cboCodigoValor.Item(varContador).Top = 150
                    Me.cboCodigoValor.Item(varContador).Left = varLeftVal
                    Me.cboValor.Item(varContador).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    Me.cboCodigoValor.Item(varContador).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    If FunFCamposObligatorios(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Or FunFCamposEditables(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Then
                        Me.cboValor.Item(varContador).Enabled = True
                    Else
                        Me.cboValor.Item(varContador).Enabled = False
                    End If
                    Me.cboValor.Item(varContador).TabIndex = varContador + varCantidadBotonesHabilitados
                    If Me.cboValor.Item(varContador).Enabled = True Then
                        Me.cboValor.Item(varContador).SetFocus
                    End If
                    
                    Me.cmdAgregarValores.Item(varContador).Visible = True
                    Me.cmdAgregarValores.Item(varContador).Top = 150
                    Me.cmdAgregarValores.Item(varContador).Left = varLeftBtn
                    Me.cmdAgregarValores.Item(varContador).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    
                    If Me.proParametroProducto.Item(varContador + 1).proEtiqueta = "Exento" Then varIndiceCboExento = varContador

                    
'                    If Me.proParametroProducto.Item(varContador + 1).proValidarRepetidos = "1" Or Me.proParametroProducto.Item(varContador + 1).proValidarRepetidos = "True" Then
'                        Me.cmdAgregarValores.Item(varContador).Enabled = True
'                    Else
'                        Me.cmdAgregarValores.Item(varContador).Enabled = False
'                    End If
                    
                    If Me.cboValor.Item(varContador).Enabled = False Then
                        Me.cmdAgregarValores.Item(varContador).Enabled = False
                    End If
                    
                    If Me.cmdAgregarValores.Item(varContador).Enabled = True And (Me.proParametroProducto.Item(varContador + 1).proValidarRepetidos = "1" Or Me.proParametroProducto.Item(varContador + 1).proValidarRepetidos = "True") Then
                        varCantidadBotonesHabilitados = 1
                        Me.cmdAgregarValores.Item(varContador).Enabled = True
                        Me.cmdAgregarValores.Item(varContador).TabIndex = varContador + varCantidadBotonesHabilitados
                    Else
                        Me.cmdAgregarValores.Item(varContador).Enabled = False
                    End If
                    
                Case "B"
                    Me.chkValor.Item(varContador).Visible = True
                    Me.chkValor.Item(varContador).Top = 150
                    Me.chkValor.Item(varContador).Left = varLeftVal
                    Me.chkValor.Item(varContador).Value = 0
                    Me.chkValor.Item(varContador).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    If FunFCamposObligatorios(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Or FunFCamposEditables(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Then
                        Me.chkValor.Item(varContador).Enabled = True
                    Else
                        Me.chkValor.Item(varContador).Enabled = False
                    End If
                        
                    Me.chkValor.Item(varContador).TabIndex = varContador
                    If Me.chkValor.Item(varContador).Enabled = True Then
                        Me.chkValor.Item(varContador).SetFocus
                    End If
                Case Else
                    MsgBox "El tipo de campo no es valido.", vbCritical, App.Title
                    Exit Sub
            End Select
            If Me.proParametroProducto.Item(varContador + 1).proTipo = "L" Then
'                If Me.proParametroProducto.Item(varContador + 1).proCampoPadre = "" Then
'                    If Not Me.proParametroProducto.Item(varContador + 1).MetConsultarValores Then
'                        MsgBox "Error al consultar los valores.", vbCritical, App.Title
'                    End If
'                End If
                For varContadorValores = 1 To Me.proParametroProducto.Item(varContador + 1).proValores.Count
                    Me.cboValor.Item(varContador).AddItem Me.proParametroProducto.Item(varContador + 1).proValores.Item(varContadorValores).proValorDesc
                    Me.cboCodigoValor.Item(varContador).AddItem Me.proParametroProducto.Item(varContador + 1).proValores.Item(varContadorValores).proValorID
                Next varContadorValores
                
                                'Asignar valores por defecto a los combos
                If Me.proParametroProducto.Item(varContador + 1).proMascara <> "" Then
                    For varIndice = 0 To cboCodigoValor(varContador).ListCount - 1
                        If cboCodigoValor(varContador).List(varIndice) = Me.proParametroProducto.Item(varContador + 1).proMascara Then
                            cboCodigoValor(varContador).ListIndex = varIndice
                            cboValor(varContador).ListIndex = varIndice
                            Exit For
                        End If
                    Next
                End If
      
            End If
        Else
            Load Me.pnlEtiqueta(varContador)
            Me.pnlEtiqueta.Item(varContador).Visible = True
            Me.pnlEtiqueta.Item(varContador).Caption = Me.proParametroProducto.Item(varContador + 1).proEtiqueta
            Me.pnlEtiqueta.Item(varContador).Left = varLeftPnl
            If varContador = 20 Then
                Me.pnlEtiqueta.Item(varContador).Top = 150
            Else
                Me.pnlEtiqueta.Item(varContador).Top = 150 + (330 * varTop)
            End If
            
            Select Case Me.proParametroProducto.Item(varContador + 1).proTipo
                Case "T"
                    varCantidad = Me.txtValor.Count
                    If Me.txtValor.Item(varCantidad - 1).Visible = False And varCantidad = 1 Then
                        Me.txtValor.Item(varCantidad - 1).Visible = True
                    Else
                        Load Me.txtValor(varCantidad)
                        
                        varCantidad = Me.txtValor.Count
                        Me.txtValor.Item(varCantidad - 1).Visible = True
                    End If
                    
                    If varContador = 20 Then
                        Me.txtValor.Item(varCantidad - 1).Top = 150
                    Else
                        Me.txtValor.Item(varCantidad - 1).Top = 150 + (330 * varTop)
                    End If
                    
                    Me.txtValor.Item(varCantidad - 1).Left = varLeftVal
                    Me.txtValor.Item(varCantidad - 1).Text = ""
                    Me.txtValor.Item(varCantidad - 1).MaxLength = Me.proParametroProducto.Item(varContador + 1).proTamaño
                    Me.txtValor.Item(varCantidad - 1).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    If FunFCamposObligatorios(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Or FunFCamposEditables(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Then
                        Me.txtValor.Item(varCantidad - 1).Enabled = True
                    Else
                        Me.txtValor.Item(varCantidad - 1).Enabled = False
                    End If
                    Me.txtValor.Item(varCantidad - 1).TabIndex = varContador
                Case "F"
                    varCantidad = Me.dtValor.Count
                    If Me.dtValor.Item(varCantidad - 1).Visible = False And varCantidad = 1 Then
                        Me.dtValor.Item(varCantidad - 1).Visible = True
                    Else
                        Load Me.dtValor(varCantidad)
                        
                        varCantidad = Me.dtValor.Count
                        Me.dtValor.Item(varCantidad - 1).Visible = True
                    End If
                    If varContador = 10 Then
                        Me.dtValor.Item(varCantidad - 1).Top = 150
                    Else
                        Me.dtValor.Item(varCantidad - 1).Top = 150 + (330 * varTop)
                    End If
                        
                    Me.dtValor.Item(varCantidad - 1).Left = varLeftVal
                    Me.dtValor.Item(varCantidad - 1).Value = Now
                    Me.dtValor.Item(varCantidad - 1).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    If FunFCamposObligatorios(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Or FunFCamposEditables(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Then
                        Me.dtValor.Item(varCantidad - 1).Enabled = True
                        Else
                        Me.dtValor.Item(varCantidad - 1).Enabled = False
                    End If
                    Me.dtValor.Item(varCantidad - 1).TabIndex = varContador
                Case "L"
                    varCantidad = Me.cboValor.Count
                    If Me.cboValor.Item(varCantidad - 1).Visible = False And varCantidad = 1 Then
                        Me.cboValor.Item(varCantidad - 1).Visible = True
                        Me.cmdAgregarValores.Item(varContador).Visible = True
                    Else
                        Load Me.cboValor(varCantidad)
                        Load Me.cboCodigoValor(varCantidad)
                        Load Me.cmdAgregarValores(varCantidad)
                        
                        varCantidad = Me.cboValor.Count
                        Me.cboValor.Item(varCantidad - 1).Visible = True
                        Me.cmdAgregarValores.Item(varCantidad - 1).Visible = True
                    End If

                    Me.cboValor.Item(varCantidad - 1).Top = 150 + (330 * varTop)
                    Me.cboValor.Item(varCantidad - 1).Left = varLeftVal
                    Me.cboValor.Item(varCantidad - 1).ListIndex = -1
                    
                    Me.cboCodigoValor.Item(varCantidad - 1).Top = 150 + (330 * varTop)
                    Me.cboCodigoValor.Item(varCantidad - 1).Left = varLeftVal
                    
                    If Me.proParametroProducto.Item(varContador + 1).proEtiqueta = "Exento" Then varIndiceCboExento = varCantidad - 1
                   
                    Me.cboValor.Item(varCantidad - 1).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    Me.cboCodigoValor.Item(varCantidad - 1).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    
                    If FunFCamposObligatorios(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Or FunFCamposEditables(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Then
                        Me.cboValor.Item(varCantidad - 1).Enabled = True
                     Else
                        Me.cboValor.Item(varCantidad - 1).Enabled = False
                     End If
                    Me.cboValor.Item(varCantidad - 1).TabIndex = varContador + varCantidadBotonesHabilitados
                    
                    Me.cmdAgregarValores.Item(varCantidad - 1).Top = 150 + (330 * varTop)
                    Me.cmdAgregarValores.Item(varCantidad - 1).Left = varLeftBtn
                    Me.cmdAgregarValores.Item(varCantidad - 1).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    
'                    If Me.proParametroProducto.Item(varContador + 1).proValidarRepetidos = "1" Or Me.proParametroProducto.Item(varContador + 1).proValidarRepetidos = "True" Then
'                        Me.cmdAgregarValores.Item(varCantidad - 1).Enabled = True
'                    Else
'                        Me.cmdAgregarValores.Item(varCantidad - 1).Enabled = False
'                    End If
'
                    If Me.cboValor.Item(varCantidad - 1).Enabled = False Then
                        Me.cmdAgregarValores.Item(varCantidad - 1).Enabled = False
                    End If
                    
                    If Me.cmdAgregarValores.Item(varCantidad - 1).Enabled = True And (Me.proParametroProducto.Item(varContador + 1).proValidarRepetidos = "1" Or Me.proParametroProducto.Item(varContador + 1).proValidarRepetidos = "True") Then
                        varCantidadBotonesHabilitados = varCantidadBotonesHabilitados + 1
                        Me.cmdAgregarValores.Item(varCantidad - 1).TabIndex = varContador + varCantidadBotonesHabilitados
                    End If
                Case "B"
                    varCantidad = Me.chkValor.Count
                    If Me.chkValor.Item(varCantidad - 1).Visible = False And varCantidad = 1 Then
                        Me.chkValor.Item(varCantidad - 1).Visible = True
                    Else
                        Load Me.chkValor(varCantidad)
                        
                        varCantidad = Me.chkValor.Count
                        Me.chkValor.Item(varCantidad - 1).Visible = True
                    End If
                    If varContador = 10 Then
                        Me.chkValor.Item(varCantidad - 1).Top = 150
                    Else
                        Me.chkValor.Item(varCantidad - 1).Top = 150 + (330 * varTop)
                    End If
                        
                    Me.chkValor.Item(varCantidad - 1).Left = varLeftVal
                    Me.chkValor.Item(varCantidad - 1).Value = 0
                    Me.chkValor.Item(varCantidad - 1).Tag = Trim(Me.proParametroProducto.Item(varContador + 1).proCampo)
                    If FunFCamposObligatorios(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Or FunFCamposEditables(Me.proParametroProducto.Item(varContador + 1).proCampo) = True Then
                        Me.chkValor.Item(varCantidad - 1).Enabled = True
                        Else
                         Me.chkValor.Item(varCantidad - 1).Enabled = False
                      End If
                                     Me.chkValor.Item(varCantidad - 1).TabIndex = varContador
                Case Else
                    MsgBox "El tipo de campo no es valido.", vbCritical, App.Title
                    Exit Sub
            End Select
            
            If Me.proParametroProducto.Item(varContador + 1).proTipo = "L" Then
'                If Me.proParametroProducto.Item(varContador + 1).proCampoPadre = "" Then
'                    If Not Me.proParametroProducto.Item(varContador + 1).MetConsultarValores Then
'                        MsgBox "Error al consultar los valores.", vbCritical, App.Title
'                    End If
'                End If
       
                For varContadorValores = 1 To Me.proParametroProducto.Item(varContador + 1).proValores.Count
                    Me.cboValor.Item(varCantidad - 1).AddItem Me.proParametroProducto.Item(varContador + 1).proValores.Item(varContadorValores).proValorDesc
                    Me.cboCodigoValor.Item(varCantidad - 1).AddItem Me.proParametroProducto.Item(varContador + 1).proValores.Item(varContadorValores).proValorID
                Next varContadorValores
                 'Asignar valores por defecto a los combos
                If Me.proParametroProducto.Item(varContador + 1).proMascara <> "" Then
                    For varIndice = 0 To cboCodigoValor(varCantidad - 1).ListCount - 1
                        If cboCodigoValor(varCantidad - 1).List(varIndice) = Me.proParametroProducto.Item(varContador + 1).proMascara Then
                            cboCodigoValor(varCantidad - 1).ListIndex = varIndice
                            cboValor(varCantidad - 1).ListIndex = varIndice
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
            
    Next varContador
    Me.cmdGuardar.TabIndex = varContador + varCantidadBotonesHabilitados + 1
    Me.cmdCancelar.TabIndex = varContador + varCantidadBotonesHabilitados + 2
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrManager
    proValor = "N"
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtValor_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    For varContador = 1 To Me.proParametroProducto.Count
        If Trim(Me.txtValor.Item(Index).Tag) = Trim(Me.proParametroProducto.Item(varContador).proCampo) Then
            Exit For
        End If
    Next varContador
    
    If Me.proParametroProducto.Item(varContador).proMascara = "A" Then
        KeyAscii = FunGLeeAlfaNumerico(KeyAscii, 1)
    Else
        KeyAscii = FunGLeeNumerico(KeyAscii)
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFCargarValores()
    Dim varContador As Integer
    Dim varContadorAux As Integer
    On Error GoTo ErrManager
    
    'Se deben recorrer cada uno de los controles para asignar el valor a las propiedades
    For varContador = 0 To Me.txtValor.Count - 1
        Select Case Me.txtValor.Item(varContador).Tag
            Case "vchUser1"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser1
            Case "vchUser2"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser2
            Case "vchUser3"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser3
            Case "vchUser4"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser4
            Case "vchUser5"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser5
            Case "vchUser6"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser6
            Case "vchUser7"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser7
            Case "vchUser8"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser8
            Case "vchUser9"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser9
            Case "vchUser10"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser10
            Case "vchUser11"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser11
            Case "vchUser12"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser12
            Case "vchUser13"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser13
            Case "vchUser14"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser14
            Case "vchUser15"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser15
            Case "vchUser16"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser16
            Case "vchUser17"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser17
            Case "vchUser18"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser18
            Case "vchUser19"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser19
            Case "vchUser20"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser20
            Case "vchUser21"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser21
            Case "vchUser22"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser22
            Case "vchUser23"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser23
            Case "vchUser24"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser24
            Case "vchUser25"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser25
            Case "vchUser26"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser26
            Case "vchUser27"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser27
            Case "vchUser28"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser28
            Case "vchUser29"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser29
            Case "vchUser30"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser30
            Case "vchUser31"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser31
            Case "vchUser32"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser32
            Case "vchUser33"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser33
            Case "vchUser34"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser34
            Case "vchUser35"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser35
            Case "vchUser36"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser36
            Case "vchUser37"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser37
            Case "vchUser38"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser38
            Case "vchUser39"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser39
            Case "vchUser40"
                Me.txtValor.Item(varContador).Text = Me.proNovedadDetalleDatosProducto.proUser40
        End Select
    Next varContador
    
    'Se deben recorrer cada uno de los controles para asignar el valor a las propiedades
    For varContador = 0 To Me.dtValor.Count - 1
        Select Case Me.dtValor.Item(varContador).Tag
            Case "vchUser1"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser1) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser1
                End If
            Case "vchUser2"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser2) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser2
                End If
            Case "vchUser3"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser3) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser3
                End If
            Case "vchUser4"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser4) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser4
                End If
            Case "vchUser5"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser5) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser5
                End If
            Case "vchUser6"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser6) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser6
                End If
            Case "vchUser7"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser7) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser7
                End If
            Case "vchUser8"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser8) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser8
                End If
            Case "vchUser9"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser9) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser9
                End If
            Case "vchUser10"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser10) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser10
                End If
            Case "vchUser11"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser11) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser11
                End If
            Case "vchUser12"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser12) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser12
                End If
            Case "vchUser13"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser13) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser13
                End If
            Case "vchUser14"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser14) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser14
                End If
            Case "vchUser15"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser15) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser15
                End If
            Case "vchUser16"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser16) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser16
                End If
            Case "vchUser17"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser17) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser17
                End If
            Case "vchUser18"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser18) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser18
                End If
            Case "vchUser19"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser19) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser19
                End If
            Case "vchUser20"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser20) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser20
                End If
            Case "vchUser21"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser21) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser21
                End If
            Case "vchUser22"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser22) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser22
                End If
            Case "vchUser23"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser23) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser23
                End If
            Case "vchUser24"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser24) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser24
                End If
            Case "vchUser25"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser25) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser25
                End If
            Case "vchUser26"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser26) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser26
                End If
            Case "vchUser27"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser27) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser27
                End If
            Case "vchUser28"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser28) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser28
                End If
            Case "vchUser29"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser29) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser29
                End If
            Case "vchUser30"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser30) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser30
                End If
            Case "vchUser31"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser31) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser31
                End If
            Case "vchUser32"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser32) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser32
                End If
            Case "vchUser33"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser33) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser33
                End If
            Case "vchUser34"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser34) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser34
                End If
            Case "vchUser35"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser35) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser35
                End If
            Case "vchUser36"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser36) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser36
                End If
            Case "vchUser37"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser37) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser37
                End If
            Case "vchUser38"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser38) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser38
                End If
            Case "vchUser39"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser39) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser39
                End If
            Case "vchUser40"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser40) = "" Then
                    Me.dtValor.Item(varContador).Value = Now
                Else
                    Me.dtValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser40
                End If
        End Select
    Next varContador
    
    'Se deben recorrer cada uno de los controles para asignar el valor a las propiedades
    For varContador = 0 To Me.chkValor.Count - 1
        Select Case Me.chkValor.Item(varContador).Tag
            Case "vchUser1"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser1) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser1
                End If
            Case "vchUser2"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser2) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser2
                End If
            Case "vchUser3"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser3) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser3
                End If
            Case "vchUser4"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser4) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser4
                End If
            Case "vchUser5"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser5) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser5
                End If
            Case "vchUser6"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser6) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser6
                End If
            Case "vchUser7"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser7) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser7
                End If
            Case "vchUser8"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser8) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser8
                End If
            Case "vchUser9"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser9) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser9
                End If
            Case "vchUser10"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser10) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser10
                End If
            Case "vchUser11"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser11) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser11
                End If
            Case "vchUser12"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser12) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser12
                End If
            Case "vchUser13"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser13) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser13
                End If
            Case "vchUser14"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser14) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser14
                End If
            Case "vchUser15"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser15) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser15
                End If
            Case "vchUser16"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser16) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser16
                End If
            Case "vchUser17"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser17) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser17
                End If
            Case "vchUser18"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser18) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser18
                End If
            Case "vchUser19"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser19) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser19
                End If
            Case "vchUser20"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser20) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser20
                End If
            Case "vchUser21"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser21) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser21
                End If
            Case "vchUser22"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser22) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser22
                End If
            Case "vchUser23"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser23) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser23
                End If
            Case "vchUser24"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser24) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser24
                End If
            Case "vchUser25"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser25) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser25
                End If
            Case "vchUser26"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser26) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser26
                End If
            Case "vchUser27"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser27) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser27
                End If
            Case "vchUser28"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser28) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser28
                End If
            Case "vchUser29"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser29) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser29
                End If
            Case "vchUser30"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser30) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser30
                End If
            Case "vchUser31"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser31) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser31
                End If
            Case "vchUser32"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser32) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser32
                End If
            Case "vchUser33"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser33) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser33
                End If
            Case "vchUser34"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser34) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser34
                End If
            Case "vchUser35"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser35) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser35
                End If
            Case "vchUser36"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser36) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser36
                End If
            Case "vchUser37"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser37) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser37
                End If
            Case "vchUser38"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser38) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser38
                End If
            Case "vchUser39"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser39) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser39
                End If
            Case "vchUser40"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser40) = "" Then
                    Me.chkValor.Item(varContador).Value = 0
                Else
                    Me.chkValor.Item(varContador).Value = Me.proNovedadDetalleDatosProducto.proUser40
                End If
        End Select
    Next varContador
    
    For varContador = 0 To Me.cboValor.Count - 1
        Select Case Me.cboValor.Item(varContador).Tag
            Case "vchUser1"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser1) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser1) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser2"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser2) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser2) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser3"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser3) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser3) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser4"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser4) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser4) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser5"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser5) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser5) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser6"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser6) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser6) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser7"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser7) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser7) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser8"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser8) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser8) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser9"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser9) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser9) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser10"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser10) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser10) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser11"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser11) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser11) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser12"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser12) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser12) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser13"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser13) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser13) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser14"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser14) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser14) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser15"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser15) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser15) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser16"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser16) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser16) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser17"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser17) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser17) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser18"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser18) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser18) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser19"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser19) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser19) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser20"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser20) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser20) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser21"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser21) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser21) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser22"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser22) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser22) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser23"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser23) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser23) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser24"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser24) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser24) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser25"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser25) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser25) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser26"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser26) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser26) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser27"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser27) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser27) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser28"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser28) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser28) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser29"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser29) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser29) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser30"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser30) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser30) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser31"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser31) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser31) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser32"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser32) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser32) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser33"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser33) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser33) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser34"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser34) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser34) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser35"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser35) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser35) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser36"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser36) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser36) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser37"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser37) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser37) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser38"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser38) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser38) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser39"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser39) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser39) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
            Case "vchUser40"
                If Trim(Me.proNovedadDetalleDatosProducto.proUser40) <> "" Then
                    For varContadorAux = 0 To Me.cboCodigoValor.Item(varContador).ListCount
                        Me.cboCodigoValor.Item(varContador).ListIndex = varContadorAux
                        If Trim(Me.cboCodigoValor.Item(varContador).Text) = Trim(Me.proNovedadDetalleDatosProducto.proUser40) Then
                            Exit For
                        End If
                    Next varContadorAux
                    Me.cboValor.Item(varContador).ListIndex = Me.cboCodigoValor.Item(varContador).ListIndex
                End If
        End Select
    Next varContador
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Public Function FunFCamposObligatorios(parCampo As String) As Boolean
    Dim varCategoria As Categoria
    Dim varContador As Integer
    On Error GoTo ErrManager
         
    varCategoria = varProceso.proIncidentCategory
    If parCampo = "" Then
        FunFCamposObligatorios = False
        Exit Function
    End If
    If varCategoria = Venta Then
        If Trim(varProceso.proOTId) = "" Then
            For varContador = 1 To Me.proParametroProducto.Count
                If Trim(parCampo) = Trim(Me.proParametroProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proParametroProducto.Item(varContador).proObligatorioVenta = True Then
                FunFCamposObligatorios = True
            Else
                FunFCamposObligatorios = False
            End If
        Else
            For varContador = 1 To Me.proParametroProducto.Count
                If Trim(parCampo) = Trim(Me.proParametroProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proParametroProducto.Item(varContador).proObligatorioOT = True Then
                FunFCamposObligatorios = True
            Else
                FunFCamposObligatorios = False
            End If
        End If
    ElseIf varCategoria = Atencion Then
        If Trim(varProceso.proOTId) = "" Then
            For varContador = 1 To Me.proParametroProducto.Count
                If Trim(parCampo) = Trim(Me.proParametroProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proParametroProducto.Item(varContador).proObligatorioAtencion = True Then
                FunFCamposObligatorios = True
            Else
                FunFCamposObligatorios = False
            End If
        Else
            For varContador = 1 To Me.proParametroProducto.Count
                If Trim(parCampo) = Trim(Me.proParametroProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proParametroProducto.Item(varContador).proObligatorioOT = True Then
                FunFCamposObligatorios = True
            Else
                FunFCamposObligatorios = False
            End If
        End If
    End If
    
        Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function FunFCamposEditables(parCampo As String) As Boolean
    Dim varCategoria As Categoria
    Dim varContador As Integer
    On Error GoTo ErrManager
        
    varCategoria = varProceso.proIncidentCategory
    If varCategoria = Venta Then
        If Trim(varProceso.proOTId) = "" Then
            For varContador = 1 To Me.proParametroProducto.Count
                If Trim(parCampo) = Trim(Me.proParametroProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proParametroProducto.Item(varContador).proEditableVenta = True Then
                FunFCamposEditables = True
            Else
                FunFCamposEditables = False
            End If
        Else
            For varContador = 1 To Me.proParametroProducto.Count
                If Trim(parCampo) = Trim(Me.proParametroProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proParametroProducto.Item(varContador).proEditableOT = True Then
                FunFCamposEditables = True
            Else
                FunFCamposEditables = False
            End If
        End If
    ElseIf varCategoria = Atencion Then
        If Trim(varProceso.proOTId) = "" Then
            For varContador = 1 To Me.proParametroProducto.Count
                If Trim(parCampo) = Trim(Me.proParametroProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proParametroProducto.Item(varContador).proEditableAtencion = True Then
                FunFCamposEditables = True
            Else
                FunFCamposEditables = False
            End If
        Else
            For varContador = 1 To Me.proParametroProducto.Count
                If Trim(parCampo) = Trim(Me.proParametroProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proParametroProducto.Item(varContador).proEditableOT = True Then
                FunFCamposEditables = True
            Else
                FunFCamposEditables = False
            End If
        End If
    End If
    
        Exit Function
ErrManager:
    SubGMuestraError
End Function


Function ValidarCamposObligatorios() As Boolean
    Dim varContador As Integer
    On Error GoTo ErrManager
    ValidarCamposObligatorios = False
    
    For varContador = 0 To Me.txtValor.Count - 1
        If FunFCamposObligatorios(txtValor(varContador).Tag) Then
            If txtValor(varContador).Enabled And txtValor(varContador).Text = "" Then
                txtValor(varContador).SetFocus
                MsgBox "Debe digitar un valor en este campo"
                Exit Function
            End If
        End If
    Next
    For varContador = 0 To Me.dtValor.Count - 1
    If FunFCamposObligatorios(dtValor(varContador).Tag) Then
        If dtValor(varContador).Enabled And dtValor(varContador).Value = "" Then
            dtValor(varContador).SetFocus
            MsgBox "Debe seleccionar un valor en este campo"
            Exit Function
        End If
        End If
    Next
    For varContador = 0 To Me.cboValor.Count - 1
    If FunFCamposObligatorios(cboValor(varContador).Tag) Then
        If cboValor(varContador).Enabled And cboValor(varContador).ListIndex < 0 Then
            cboValor(varContador).SetFocus
            MsgBox "Debe seleccionar un valor en este campo"
            Exit Function
        End If
        End If
    Next
    
    
'    For varContador = 0 To Me.txtValor.Count - 1
'
'        If txtValor(varContador).Enabled And txtValor(varContador).Text = "" Then
'            txtValor(varContador).SetFocus
'            MsgBox "Debe digitar un valor en este campo"
'            Exit Function
'        End If
'    Next
'    For varContador = 0 To Me.dtValor.Count - 1
'        If dtValor(varContador).Enabled And dtValor(varContador).Value = "" Then
'            dtValor(varContador).SetFocus
'            MsgBox "Debe seleccionar un valor en este campo"
'            Exit Function
'        End If
'    Next
'    For varContador = 0 To Me.cboValor.Count - 1
'        If cboValor(varContador).Enabled And cboValor(varContador).ListIndex < 0 Then
'            cboValor(varContador).SetFocus
'            MsgBox "Debe seleccionar un valor en este campo"
'            Exit Function
'        End If
'    Next
    ValidarCamposObligatorios = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function
