VERSION 5.00
Begin VB.Form frmCambioTipoLinea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio Tipo Linea"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2490
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ListBox lstTiposLinea 
      Height          =   645
      ItemData        =   "frmCambioTipoLinea.frx":0000
      Left            =   120
      List            =   "frmCambioTipoLinea.frx":0002
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton CmdCambiarTipoLinea 
      Caption         =   "&Cambiar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblInstrucciones 
      Caption         =   "Seleccione el tipo de linea destino"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmCambioTipoLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'   DESCRIPCION         : En este fomulario se lista los tipos delinea en uso y edicion para que
'                         se selecionen y retorna el selecionado .
'   PARAMETROS          :
'                         proDatosProducto       ClaDatosproducto
'
'   RETORNO             :
'                         proTipoLinea           Entero(id tipo linea)
'
'   EJEMPLO             :
'                           Set frmCambioTipoLinea.proDatosProducto = proDatosProducto
'                           frmCambioTipoLinea.Show vbModal
'                           varNuevoTipoLinea = frmCambioTipoLinea.proTipoLinea
'
'*************************************************************************************************
'   MODIFICADO POR      : Carlos Leonardo Villamil (I&T)
'   DESCRIPCION CAMBIO  : El cambio permite cambiar la asociacion entre tipos de linea en curso o
'                         en uso y numeros publicos en uso.
'   VERCION             : 3.7.4
'   FECHA               : 09-JUL-09
'*************************************************************************************************

Private proConexions As ADODB.Connection
Private proDatosProductos As claDatosProducto
Private ProTipoLineas As Long
Private Const conSinSeleccion As Integer = -1
Private Const conTipoLineabasica As String = 1
Private Const conTipoLineaE1 As String = 2
Private Const conTipoTroncalSip As String = 76668
Private Const conTipoVirtual As String = 83806


Public Property Set ProConnection(parConexion As ADODB.Connection)
        Set proConexions = parConexion
End Property

Public Property Set proDatosProducto(ByVal parDatosProductos As claDatosProducto)
            Set proDatosProductos = parDatosProductos
            proDatosProductos.MetConsultarNovedadDetalleDatosProducto
            proDatosProductos.MetConsultarDetalles
End Property

Public Property Get proTipoLinea() As Long
        proTipoLinea = ProTipoLineas
End Property


Private Sub CmdCambiarTipoLinea_Click()
If lstTiposLinea.Text = "" Then
    ProTipoLineas = -1
    Unload Me
Else
    ProTipoLineas = lstTiposLinea.ItemData(lstTiposLinea.ListIndex)
    Unload Me
End If
End Sub

Private Sub CmdCancelar_Click()
    ProTipoLineas = -1
    Unload Me
End Sub

Private Sub Form_Load()
        Dim varDetalleDatosProducto As claDetalleDatosProducto
        Dim varNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
        Dim var As EDCAdminVoz.claValor
        Dim varindicenovedad As Integer
        Dim varexiste  As Boolean
        Dim varTipoDeLineaDesc As String
        ProTipoLineas = conSinSeleccion
        If Not proDatosProductos.proDetalleDatosProducto Is Nothing Then
            If proDatosProductos.proDetalleDatosProducto.Count > 0 Then
                For Each varDetalleDatosProducto In proDatosProductos.proDetalleDatosProducto
                     varindicenovedad = 1
                    varexiste = False
                    While varindicenovedad <= proDatosProductos.proNovedadDetalleDatosProducto.Count And varexiste = False
                        If proDatosProductos.proNovedadDetalleDatosProducto.Item(varindicenovedad).proDetalleDatosProductoId = varDetalleDatosProducto.proDetalleDatosProductoId Then
                            varexiste = True
                        End If
                        varindicenovedad = varindicenovedad + 1
                    Wend
                    If varexiste = False Then
                        Select Case varDetalleDatosProducto.proUser1
                            Case conTipoLineabasica
                                varTipoDeLineaDesc = "B"
                            Case conTipoLineaE1
                                varTipoDeLineaDesc = "E1"
                            Case conTipoTroncalSip
                                varTipoDeLineaDesc = "IP"
                            Case conTipoVirtual
                                 varTipoDeLineaDesc = "V"
                        Case Else
                                varTipoDeLineaDesc = "B"
                        End Select
                        lstTiposLinea.AddItem (varTipoDeLineaDesc + "(" + varDetalleDatosProducto.proDetalleDatosProductoId + ")")
                        lstTiposLinea.ItemData(lstTiposLinea.ListCount - 1) = varDetalleDatosProducto.proDetalleDatosProductoId
                    End If
                Next varDetalleDatosProducto
            End If
        End If
        varindicenovedad = 1
        While varindicenovedad <= proDatosProductos.proNovedadDetalleDatosProducto.Count
            If proDatosProductos.proNovedadDetalleDatosProducto.Item(varindicenovedad).proTipoNovedadId = 1 Then
                    Select Case proDatosProductos.proNovedadDetalleDatosProducto.Item(varindicenovedad).proUser1
                        Case conTipoLineabasica
                            varTipoDeLineaDesc = "B"
                        Case conTipoLineaE1
                            varTipoDeLineaDesc = "E1"
                        Case conTipoTroncalSip
                            varTipoDeLineaDesc = "IP"
                        Case conTipoVirtual
                             varTipoDeLineaDesc = "V"
                    Case Else
                            varTipoDeLineaDesc = "B"
                    End Select
                   lstTiposLinea.AddItem (varTipoDeLineaDesc + "(" + proDatosProductos.proNovedadDetalleDatosProducto.Item(varindicenovedad).proNovedadDetalleDatosProductoId + ")")
                   lstTiposLinea.ItemData(lstTiposLinea.ListCount - 1) = proDatosProductos.proNovedadDetalleDatosProducto.Item(varindicenovedad).proNovedadDetalleDatosProductoId
            End If
            varindicenovedad = varindicenovedad + 1
        Wend
End Sub
