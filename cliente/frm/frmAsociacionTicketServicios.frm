VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAsociacionTicketServicios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asociación de Servicios"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   2655
      Width           =   6240
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Líneas no seleccionadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   420
         TabIndex        =   6
         Top             =   180
         Width           =   1800
      End
      Begin VB.Label lblInsertar 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   165
         Left            =   150
         TabIndex        =   5
         Top             =   210
         Width           =   165
      End
      Begin VB.Label lblSeleccionModificacion 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   165
         Left            =   2910
         TabIndex        =   4
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Líneas seleccionadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3180
         TabIndex        =   3
         Top             =   210
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdProducto 
      Height          =   2535
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin VB.CommandButton btnGuardar 
      Caption         =   "&Guardar"
      Default         =   -1  'True
      Height          =   345
      Left            =   6435
      TabIndex        =   1
      Top             =   2790
      Width           =   1425
   End
End
Attribute VB_Name = "frmAsociacionTicketServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public proDatosProducto As claDatosProducto
Public proIncidenteId As Long
Public proclaTicketxDetalleDP As claticketxdetalledatosproducto
Public proConexion As ADODB.Connection
Public proOnyx As EDCVoz.claONYX
Public provchSerialNumber As String

Private Sub btnGuardar_Click()
        proclaTicketxDetalleDP.proiIncidentId = Me.proIncidenteId
        Call proclaTicketxDetalleDP.FunGEliminar
        For i = 1 To Me.grdProducto.Rows - 1
            grdProducto.Row = i
            grdProducto.Col = 1
            If grdProducto.CellBackColor = Me.lblSeleccionModificacion.BackColor Then
                proclaTicketxDetalleDP.proiDatosProductoId = Me.proDatosProducto.proDetalleDatosProducto.Item(i).proDatosProductoId
                proclaTicketxDetalleDP.proiDetalleDatosProductoId = proDatosProducto.proDetalleDatosProducto.Item(i).proDetalleDatosProductoId
                Call proclaTicketxDetalleDP.FunGGuardar
            End If
        Next

End Sub

Private Sub Form_Load()
            Set proclaTicketxDetalleDP = New claticketxdetalledatosproducto
            Set proclaTicketxDetalleDP.proConexion = Me.proConexion
            If Me.proDatosProducto.MetConsultarDetalles Then
                Call SubFPintarGridDetalles
            Else
                MsgBox "Error al consultar los detalles del Tab de Datos por Servicios.", vbCritical, App.Title
                Exit Sub
            End If
End Sub

Private Sub SubFPintarGridDetalles()
    Dim varContador As Integer
    Dim varValor As String
    Dim varValorLista As EDCAdminVoz.claValor
    Dim varcolticketxdetalledatosproducto As colticketxdetalledatosproducto
    Dim varContadorAux As Integer, i As Integer, j As Integer
    Dim varValorCampo As String
    Dim bHaySelección As Boolean
    On Error GoTo ErrManager
    Set varcolticketxdetalledatosproducto = New colticketxdetalledatosproducto
    Set varcolticketxdetalledatosproducto.proConexion = Me.proConexion
    varcolticketxdetalledatosproducto.proiIncidentId = Me.proIncidenteId
    If varcolticketxdetalledatosproducto.FunGConsulta Then bHaySelección = True
    grdProducto.Clear
    If Me.proDatosProducto.proParametrosProducto Is Nothing Then
        Exit Sub
    End If
    
    If Me.proDatosProducto.proParametrosProducto.Count = 0 Then
        Exit Sub
    End If
    Call SubFInicializarGridDetalle
    varValor = ""
    Me.grdProducto.Rows = 1
    For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
        varValor = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId & vbTab & _
                   Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId

        For varContadorAux = 1 To Me.proDatosProducto.proParametrosProducto.Count
            Select Case Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proTipo
                Case "L"
                    Set varValorLista = New EDCAdminVoz.claValor
                    Set varValorLista.proConexion = Me.proConexion
                
                    Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proCampo)
                        Case "vchUser1"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser1
                        Case "vchUser2"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser2
                        Case "vchUser3"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser3
                        Case "vchUser4"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser4
                        Case "vchUser5"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser5
                        Case "vchUser6"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser6
                        Case "vchUser7"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser7
                        Case "vchUser8"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser8
                        Case "vchUser9"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser9
                        Case "vchUser10"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser10
                        Case "vchUser11"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser11
                        Case "vchUser12"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser12
                        Case "vchUser13"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser13
                        Case "vchUser14"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser14
                        Case "vchUser15"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser15
                        Case "vchUser16"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser16
                        Case "vchUser17"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser17
                        Case "vchUser18"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser18
                        Case "vchUser19"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser19
                        Case "vchUser20"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser20
                        Case "vchUser21"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser21
                        Case "vchUser22"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser22
                        Case "vchUser23"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser23
                        Case "vchUser24"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser24
                        Case "vchUser25"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser25
                        Case "vchUser26"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser26
                        Case "vchUser27"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser27
                        Case "vchUser28"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser28
                        Case "vchUser29"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser29
                        Case "vchUser30"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser30
                        Case "vchUser31"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser31
                        Case "vchUser32"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser32
                        Case "vchUser33"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser33
                        Case "vchUser34"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser34
                        Case "vchUser35"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser35
                        Case "vchUser36"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser36
                        Case "vchUser37"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser37
                        Case "vchUser38"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser38
                        Case "vchUser39"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser39
                        Case "vchUser40"
                            varValorLista.proValorId = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser40
                    End Select
                    
                    If varValorLista.MetConsultar Then
                        varValor = varValor & vbTab & varValorLista.proValorDesc
                        Set varValorLista = Nothing
                    Else
                        MsgBox "Error al consultar el valor.", vbCritical, App.Title
                        Exit Sub
                    End If
                    
                Case "B"
                    Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proCampo)
                        Case "vchUser1"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser1 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser2"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser2 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser3"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser3 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser4"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser4 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser5"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser5 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser6"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser6 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser7"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser7 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser8"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser8 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser9"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser9 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser10"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser10 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser11"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser11 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser12"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser12 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser13"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser13 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser14"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser14 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser15"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser15 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser16"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser16 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser17"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser17 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser18"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser18 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser19"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser19 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser20"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser20 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser21"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser21 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser22"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser22 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser23"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser23 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser24"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser24 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser25"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser25 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser26"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser26 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser27"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser27 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser28"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser28 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser29"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser29 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser30"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser30 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser31"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser31 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser32"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser32 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser33"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser33 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser34"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser34 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser35"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser35 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser36"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser36 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser37"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser37 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser38"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser38 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser39"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser39 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                        Case "vchUser40"
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser40 = 1 Then
                                varValorCampo = "SI"
                            Else
                                varValorCampo = "NO"
                            End If
                    End Select
                
                    varValor = varValor & vbTab & varValorCampo
                Case Else
                
                    Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux).proCampo)
                        Case "vchUser1"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser1
                        Case "vchUser2"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser2
                        Case "vchUser3"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser3
                        Case "vchUser4"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser4
                        Case "vchUser5"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser5
                        Case "vchUser6"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser6
                        Case "vchUser7"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser7
                        Case "vchUser8"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser8
                        Case "vchUser9"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser9
                        Case "vchUser10"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser10
                        Case "vchUser11"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser11
                        Case "vchUser12"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser12
                        Case "vchUser13"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser13
                        Case "vchUser14"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser14
                        Case "vchUser15"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser15
                        Case "vchUser16"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser16
                        Case "vchUser17"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser17
                        Case "vchUser18"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser18
                        Case "vchUser19"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser19
                        Case "vchUser20"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser20
                        Case "vchUser21"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser21
                        Case "vchUser22"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser22
                        Case "vchUser23"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser23
                        Case "vchUser24"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser24
                        Case "vchUser25"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser25
                        Case "vchUser26"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser26
                        Case "vchUser27"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser27
                        Case "vchUser28"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser28
                        Case "vchUser29"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser29
                        Case "vchUser30"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser30
                        Case "vchUser31"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser31
                        Case "vchUser32"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser32
                        Case "vchUser33"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser33
                        Case "vchUser34"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser34
                        Case "vchUser35"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser35
                        Case "vchUser36"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser36
                        Case "vchUser37"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser37
                        Case "vchUser38"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser38
                        Case "vchUser39"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser39
                        Case "vchUser40"
                            varValorCampo = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser40
                    End Select
                
                    varValor = varValor & vbTab & varValorCampo
            End Select
        Next varContadorAux
        
        Me.grdProducto.AddItem varValor
        
        'No mostrar los regsitros eliminados
        If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proRecordStatus = 0 Then
            Me.grdProducto.RowHeight(Me.grdProducto.Rows - 1) = 0
        End If
        
        'No mostrar los registros cancelados
        If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId = "C" Then
            Me.grdProducto.RowHeight(Me.grdProducto.Rows - 1) = 0
        Else
            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proStatusId = "C" Then
                Call SubFPintarFila(Me.grdProducto, Me.grdProducto.Rows - 1, Me.lblSeleccionModificacion.BackColor)
            End If
        End If
    Next
    ' aplicar colores a la filas seleccionadas
    If bHaySelección Then
        For i = 1 To Me.grdProducto.Rows - 1
            grdProducto.Row = i
            For j = 1 To varcolticketxdetalledatosproducto.Count
                If proDatosProducto.proDetalleDatosProducto.Item(i).proDetalleDatosProductoId = varcolticketxdetalledatosproducto.Item(j).proiDetalleDatosProductoId Then
                    Call SubFPintarFila(Me.grdProducto, i, Me.lblSeleccionModificacion.BackColor)
                End If
            Next
        Next
    End If
    grdProducto.Cols = 10
    Exit Sub
ErrManager:
    SubGMuestraError

End Sub

Private Sub SubFInicializarGridDetalle()
    Dim varContador As Integer
    On Error GoTo ErrManager
        
        If Me.proDatosProducto.proParametrosProducto Is Nothing Then
            Exit Sub
        End If
        
        If Me.proDatosProducto.proParametrosProducto.Count = 0 Then
            MsgBox "El producto del incidente seleccionado no tiene campos parametrizados."
            Exit Sub
        Else
       
            With Me.grdProducto
                .Cols = Me.proDatosProducto.proParametrosProducto.Count + 2
                .Rows = 1
                .Row = 0
                
                .Col = 0
                .CellAlignment = 4
                .ColWidth(0) = 0
                .TextMatrix(0, 0) = "Código"
                
                .Col = 1
                .CellAlignment = 4
                .ColWidth(1) = 0
                .TextMatrix(0, 1) = "Codigo Estado"
                
                For varContador = 1 To Me.proDatosProducto.proParametrosProducto.Count
                    .Col = varContador
                    .CellAlignment = 4
                    If Me.proDatosProducto.proParametrosProducto.Item(varContador).proTipo = "F" Then
                        .ColWidth(varContador + 1) = 2000
                    Else
                        .ColWidth(varContador + 1) = 1500
                    End If
                    .TextMatrix(0, varContador + 1) = Me.proDatosProducto.proParametrosProducto.Item(varContador).proEtiqueta
                Next varContador
                
                .Col = 0
                .Row = 0
            End With
        End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub grdProducto_DblClick()
    If Me.proDatosProducto.proDetalleDatosProducto.Item(Me.grdProducto.Row).proStatusId = "A" Then
        If Me.proDatosProducto.proDetalleDatosProducto.Item(Me.grdProducto.Row).proSeleccion = "0" Then
            Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados + 1
            Me.proDatosProducto.proDetalleDatosProducto.Item(Me.grdProducto.Row).proSeleccion = 1
            Call SubFPintarFila(Me.grdProducto, Me.grdProducto.Row, Me.lblSeleccionModificacion.BackColor)
        Else
            Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados - 1
            Me.proDatosProducto.proDetalleDatosProducto.Item(Me.grdProducto.Row).proSeleccion = 0
            Call SubFPintarFila(Me.grdProducto, Me.grdProducto.Row, Me.lblInsertar.BackColor)
        End If
    End If
End Sub
