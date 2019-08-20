VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAprobacionNumeros 
   Caption         =   "Aprobación de números por clasificación"
   ClientHeight    =   9330
   ClientLeft      =   2325
   ClientTop       =   -2595
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   10020
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Convenciones"
      Height          =   1215
      Left            =   120
      TabIndex        =   24
      Top             =   11160
      Width           =   1815
      Begin VB.Label lb 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   150
         Index           =   7
         Left            =   120
         TabIndex        =   32
         Top             =   260
         Width           =   150
      End
      Begin VB.Label lb 
         BackColor       =   &H00D6D0B8&
         BorderStyle     =   1  'Fixed Single
         Height          =   150
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   500
         Width           =   150
      End
      Begin VB.Label lbAprobado 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   150
         Index           =   5
         Left            =   120
         TabIndex        =   30
         Top             =   700
         Width           =   150
      End
      Begin VB.Label lbRechazado 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   150
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   950
         Width           =   150
      End
      Begin VB.Label Label9 
         Caption         =   "No seleccionado"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Seleccionado"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   500
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Aprobado"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   700
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Rechazado"
         Height          =   225
         Left            =   360
         TabIndex        =   25
         Top             =   950
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Instrucciones"
      Height          =   1215
      Left            =   2040
      TabIndex        =   22
      Top             =   11160
      Width           =   11955
      Begin VB.Label Label10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   10125
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   19500
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pendientes"
      TabPicture(0)   =   "frmAprobacionNumeros.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSPanel1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Procesados"
      TabPicture(1)   =   "frmAprobacionNumeros.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "spnlFondo"
      Tab(1).Control(1)=   "Label2"
      Tab(1).ControlCount=   2
      Begin Threed.SSPanel spnlFondo 
         Height          =   12255
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   13725
         _Version        =   65536
         _ExtentX        =   24209
         _ExtentY        =   21616
         _StockProps     =   15
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtBuscaCampo 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   280
            Left            =   5550
            TabIndex        =   4
            Text            =   "Click aquí para Filtrar Números"
            Top             =   650
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.TextBox txtBuscar 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   280
            Left            =   120
            TabIndex        =   3
            Text            =   "Click aquí para Filtrar Clientes"
            Top             =   650
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.TextBox txtBuscaSer 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   280
            Left            =   2820
            TabIndex        =   2
            Text            =   "Click aquí para Filtrar Incidentes"
            Top             =   650
            Visible         =   0   'False
            Width           =   2610
         End
         Begin MSFlexGridLib.MSFlexGrid grdProcesados 
            Height          =   10065
            Left            =   30
            TabIndex        =   5
            Top             =   960
            Width           =   13485
            _ExtentX        =   23786
            _ExtentY        =   17754
            _Version        =   393216
            Cols            =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSForms.CommandButton CmdDeshacer 
            Height          =   405
            Left            =   12360
            TabIndex        =   35
            Top             =   120
            Width           =   1335
            Caption         =   "Deshacer"
            Size            =   "2355;714"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label lblNormal 
            BackColor       =   &H80000005&
            Height          =   315
            Left            =   6240
            TabIndex        =   9
            Top             =   660
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblaEliminar 
            BackColor       =   &H00F9D7A8&
            Height          =   315
            Left            =   5880
            TabIndex        =   8
            Top             =   660
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblColorSeleccion 
            BackColor       =   &H00D6D0B8&
            Height          =   315
            Left            =   5880
            TabIndex        =   7
            Top             =   -30
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblColorNO 
            BackColor       =   &H00BFDDC8&
            Height          =   315
            Left            =   6360
            TabIndex        =   6
            Top             =   -30
            Visible         =   0   'False
            Width           =   315
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   10575
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   13725
         _Version        =   65536
         _ExtentX        =   24209
         _ExtentY        =   18653
         _StockProps     =   15
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton cmdAprobar 
            Caption         =   "&Aprobar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   12240
            TabIndex        =   37
            Top             =   0
            Width           =   1425
         End
         Begin VB.CommandButton cmdRechazar 
            Caption         =   "&Rechazar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   12240
            TabIndex        =   36
            Top             =   480
            Width           =   1425
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Filtrar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3240
            TabIndex        =   15
            Top             =   180
            Width           =   1365
         End
         Begin VB.TextBox txtiCompanyId 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1500
            TabIndex        =   14
            Top             =   210
            Width           =   1635
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   280
            Left            =   2820
            TabIndex        =   13
            Text            =   "Click aquí para Filtrar Incidentes"
            Top             =   650
            Visible         =   0   'False
            Width           =   2610
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   280
            Left            =   120
            TabIndex        =   12
            Text            =   "Click aquí para Filtrar Clientes"
            Top             =   650
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   280
            Left            =   5550
            TabIndex        =   11
            Text            =   "Click aquí para Filtrar Números"
            Top             =   650
            Visible         =   0   'False
            Width           =   2640
         End
         Begin MSFlexGridLib.MSFlexGrid grdPendientes 
            Height          =   9465
            Left            =   30
            TabIndex        =   16
            Top             =   960
            Width           =   13485
            _ExtentX        =   23786
            _ExtentY        =   16695
            _Version        =   393216
            Cols            =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label11 
            Caption         =   "Código del Cliente:"
            Height          =   285
            Left            =   150
            TabIndex        =   21
            Top             =   300
            Width           =   1365
         End
         Begin VB.Label Label12 
            BackColor       =   &H00BFDDC8&
            Height          =   315
            Left            =   6360
            TabIndex        =   20
            Top             =   -30
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label13 
            BackColor       =   &H00D6D0B8&
            Height          =   315
            Left            =   5880
            TabIndex        =   19
            Top             =   -30
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label14 
            BackColor       =   &H00F9D7A8&
            Height          =   315
            Left            =   5880
            TabIndex        =   18
            Top             =   660
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000005&
            Height          =   315
            Left            =   6240
            TabIndex        =   17
            Top             =   660
            Visible         =   0   'False
            Width           =   315
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Procesados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   -67920
         TabIndex        =   34
         Top             =   20
         Width           =   6855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   33
         Top             =   20
         Width           =   6855
      End
   End
   Begin VB.Menu mnuPopup1 
      Caption         =   "Popup1"
      Begin VB.Menu mnuAprobar 
         Caption         =   "&Aprobar"
      End
      Begin VB.Menu mnuRechazar 
         Caption         =   "&Rechazar"
      End
   End
End
Attribute VB_Name = "frmAprobacionNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************
'   Formulario para la aprobación de números
'   Autor: Leonardo Grimaldi Salcedo
'   Modificación: Corrección de la selección y rechazo de números, opción deshacer y documentación
'   Fecha: 20/09/2006
'   Fecha 16-Nov-2006, -Se cambió la función de Aprobación/ERchazo a tabla temporal para ejecutar la
'                      Aprobación/Rechazo al cerrar la aplicación
'                      - Se corrigió el error al rechazar el número.
'******************************************************************************************************

Option Explicit

'Conexion
Public proConexion As ADODB.Connection
' Usuario de
Public proUserOnyx As String

Public VarColNumero As EDCAdminVoz.colNumero
Public VarColNumerosProcesados As EDCAdminVoz.colNumero

Public VarClasNumero As EDCAdminVoz.claNumero
Public VarClaNovedadNumero As claNovedadNumero

Public VarColTempNumeros As colTemporalAprobacionNumeros

'Variables para manejo de rangos de celda
Private varFShift As Integer
Private varFPosicion As Integer
Private varFPosicionFinal As Integer

Private Sub CmdDeshacer_Click()
On Error GoTo ErrManager
    ' Si no hay números
    If Me.grdProcesados.Rows = 1 Then
        MsgBox "No existen Registros para Deshacer.", vbInformation, App.Title
        Exit Sub
    End If
    ' Si no hay números seleccionados
    If Me.grdProcesados.Row = 0 Then
        MsgBox "Debe seleccionar el(los) número(s) que desea deshacer los cambios.", vbInformation, App.Title
        Exit Sub
    End If
    If (MsgBox("Esta seguro de querer devolver el estado este(os) número(os)?. Este(os) número(s) quedarán en Estado Pendiente.", vbYesNo + vbInformation, App.Title) = vbYes) Then
        Call Deshacer
    End If
    Exit Sub
ErrManager:
    SubGMuestraError

End Sub


'******************************************************************************************************
'                                           EVENTOS
'******************************************************************************************************

'******************************************************************************************************
'   Evento de carga del formulario
'*****************************************************************************************************R
Private Sub Form_Load()
On Error GoTo ErrManager
    ' Ocultamos los menúes emergentes
    mnuPopup1.Visible = False
    ' Inicializamos la colección de números
    Set VarColNumero = New EDCAdminVoz.colNumero
    Set VarColNumerosProcesados = New EDCAdminVoz.colNumero
    'Set VarColTempNumeros = New colTemporalAprobacionNumeros
    ' Consultamos los números pendientes de aprobación
    Call SubConsultarNumeros
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Evento de clic sobre el botón de filtrar
'******************************************************************************************************
Private Sub cmdFiltrar_Click()
On Error GoTo ErrManager
    ' Realizamos la consulta de números
    Call SubConsultarNumeros
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Evento de clic sobre el botón de rechazar
'*****************************************************************************************************R
Private Sub cmdRechazar_Click()
On Error GoTo ErrManager
    ' Verificamos que existan números para rechazar
    If Me.grdPendientes.Row = 0 Then
        MsgBox "Debe seleccionar el número que desea rechazar.", vbInformation, App.Title
        Exit Sub
    End If
    'Si no aprueba lo que se debe hacer es liberar esos registros de la tabla CT_NOVEDADNUMEROS
    'y cambiar el estado de los números a Libre "L"
    If (MsgBox("Esta seguro de querer rechazar este(os) número(os)?. Este(os) número(s) serán liberados.", vbYesNo + vbInformation, App.Title) = vbYes) Then
        Call Rechazar
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Evento de clic sobre el botón de rechazar
'*****************************************************************************************************R
Private Sub cmdAprobar_Click()
On Error GoTo ErrManager
    ' Si no hay números
    If Me.grdPendientes.Rows = 1 Then
        MsgBox "No existen campos para modificar.", vbInformation, App.Title
        Exit Sub
    End If
    ' Si no hay números seleccionados
    If Me.grdPendientes.Row = 0 Then
        MsgBox "Debe seleccionar el(los) número(s) que desea aprobar.", vbInformation, App.Title
        Exit Sub
    End If
    If (MsgBox("Esta seguro de querer aprobar este(os) número(os)?. Este(os) número(s) serán reservados.", vbYesNo + vbInformation, App.Title) = vbYes) Then
        Call Aprobar
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Evento Salir
'*****************************************************************************************************R
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrManager
Dim Opcion As String
If Me.VarColTempNumeros.Count > 0 Then
    If (MsgBox("Esta seguro de querer Aprobar/Rechazar definitivamente este(os) número(os)?", vbYesNo + vbInformation, App.Title) = vbYes) Then
        Opcion = "1"
    Else
        Opcion = "0"
    End If
    Call AprobacionFinal(Opcion)
      
    Unload Me
End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Método de cambio de selección en la grilla
'******************************************************************************************************
Private Sub grdPendientes_SelChange()
    Dim varPosicion1 As Integer
    Dim varPosicion2 As Integer
    Dim varCuenta As Integer
    Dim varCuentaColumna As Integer
    Dim varBandera As Integer

On Error GoTo ErrManager
    Me.grdPendientes.Redraw = False
    varFPosicion = Me.grdPendientes.RowSel
    varFPosicionFinal = Me.grdPendientes.Row
    'Si ya esta seleccionado el campo lo debe quitar de la selección si se ha seleccionado con control
    
    If varFPosicion = 0 Then varFPosicion = 1
    If varFPosicionFinal = 0 Then varFPosicionFinal = 1

    If varFPosicion > varFPosicionFinal Then
        varPosicion1 = varFPosicionFinal
        varPosicion2 = varFPosicion
    Else
        varPosicion1 = varFPosicion
        varPosicion2 = varFPosicionFinal
    End If

    'Si la tecla es shift, selecciona únicamente el rango indicado.
    If varFShift = 1 Or varFShift = 0 Then
        'Debe borrar lo demás
        subFLimpiarSeleccion
        varBandera = 0
    'Si la tecla es ctrl agrega a la selección anterior
    ElseIf varFShift = 2 Then
        varBandera = 1
    'Si la tecla es shift + ctrl agrega el rango a lo seleccionado
    ElseIf varFShift = 3 Then
        varBandera = 2
    Else
        subFLimpiarSeleccion
        varPosicion2 = varPosicion1
    End If

    If varFShift = 2 Or varFShift = 1 Or varFShift = 0 Then
        If varPosicion1 <> varPosicion2 Then
            Me.VarColNumero.proSeleccionados = 0
            For varCuenta = varPosicion1 To varPosicion2
                Me.VarColNumero.Item(varCuenta).proSeleccionado = "S"
                Me.grdPendientes.Row = varCuenta
                For varCuentaColumna = 0 To Me.grdPendientes.Cols - 1
                    Me.grdPendientes.Col = varCuentaColumna
                    Me.grdPendientes.CellBackColor = Me.lblaEliminar.BackColor
                Next varCuentaColumna
            Next varCuenta
        Else
                For varCuenta = varPosicion1 To varPosicion2
                    If Me.VarColNumero.Item(varCuenta).proSeleccionado = "S" Then
                       Me.VarColNumero.Item(varCuenta).proSeleccionado = "N"
                    Else
                        Me.VarColNumero.Item(varCuenta).proSeleccionado = "S"
                    End If

                    Me.grdPendientes.Row = varCuenta

                    If Me.VarColNumero.Item(varCuenta).proSeleccionado = "S" Then
                        For varCuentaColumna = 0 To Me.grdPendientes.Cols - 1
                            Me.grdPendientes.Col = varCuentaColumna
                            Me.grdPendientes.CellBackColor = Me.lblaEliminar.BackColor
                        Next varCuentaColumna
                    Else
                        For varCuentaColumna = 0 To Me.grdPendientes.Cols - 1
                            If varCuentaColumna <> 6 Then
                                Me.grdPendientes.Col = varCuentaColumna
                                Me.grdPendientes.CellBackColor = Me.lblaEliminar.BackColor
                            End If
                        Next varCuentaColumna
                    End If
                Next varCuenta
        End If
    ElseIf varFShift = 3 Then 'Shift y control
        For varCuenta = varPosicion1 To varPosicion2
            Me.VarColNumero.Item(varCuenta).proSeleccionado = "S"

            Me.grdPendientes.Row = varCuenta

            For varCuentaColumna = 0 To Me.grdPendientes.Cols - 1
                Me.grdPendientes.Col = varCuentaColumna
                Me.grdPendientes.CellBackColor = Me.lblaEliminar.BackColor
            Next varCuentaColumna
        Next varCuenta
    End If

    Me.VarColNumero.proSeleccionados = 0
    For varCuenta = 1 To Me.VarColNumero.Count
        If Me.VarColNumero.Item(varCuenta).proSeleccionado = "S" Then
            Me.VarColNumero.proSeleccionados = Me.VarColNumero.proSeleccionados + 1
        End If
    Next varCuenta

    Me.grdPendientes.Redraw = True
    'Me.grdCamposXEstado.Row = 0
    Me.grdPendientes.Row = varFPosicionFinal

    Exit Sub
ErrManager:
    SubGMuestraError
    Me.grdPendientes.Redraw = True
End Sub
'******************************************************************************************************
'                                           Métodos
'******************************************************************************************************

'******************************************************************************************************
'   Método de inicialización de las grillas
'******************************************************************************************************
Sub SubFInicializarGrid()
    On Error GoTo ErrManager
        
        With Me.grdPendientes
            .Row = 0
            .Col = 0
            .CellAlignment = 4
            .ColWidth(0) = 0
            .TextMatrix(0, 0) = "CodRegion"
        
            .Col = 1
            .CellAlignment = 4
            .ColWidth(1) = 1200
            .TextMatrix(0, 1) = "Ciudad"
            
            .Col = 2
            .CellAlignment = 4
            .ColWidth(2) = 1500
            .TextMatrix(0, 2) = "Número"
            
            .Col = 3
            .CellAlignment = 0
            .ColWidth(3) = 0
            .TextMatrix(0, 3) = "CodEstado"
    
            
            .Col = 4
            .CellAlignment = 0
            .ColWidth(4) = 0
            .TextMatrix(0, 4) = "Estado"
            
            .Col = 5
            .CellAlignment = 4
            .ColWidth(5) = 3400
            .TextMatrix(0, 5) = "Clasificación"
            
            .Col = 6
            .CellAlignment = 4
            .ColWidth(6) = 800
            .TextMatrix(0, 6) = "Incident"
            
            .Col = 7
            .CellAlignment = 4
            .ColWidth(7) = 1000
            .TextMatrix(0, 7) = "CodCliente"
            
            .Col = 8
            .CellAlignment = 4
            .ColWidth(8) = 2400
            .TextMatrix(0, 8) = "Cliente"
            
            .Rows = 1
        End With
    Me.grdProcesados.Clear
    Me.grdProcesados.Rows = 1
    With Me.grdProcesados
            .Row = 0
            .Col = 0
            .CellAlignment = 4
            .ColWidth(0) = 0
            .TextMatrix(0, 0) = "CodRegion"
        
            .Col = 1
            .CellAlignment = 4
            .ColWidth(1) = 1200
            .TextMatrix(0, 1) = "Ciudad"
            
            .Col = 2
            .CellAlignment = 4
            .ColWidth(2) = 1500
            .TextMatrix(0, 2) = "Número"
            
            .Col = 3
            .CellAlignment = 0
            .ColWidth(3) = 0
            .TextMatrix(0, 3) = "CodEstado"
    
            
            .Col = 4
            .CellAlignment = 0
            .ColWidth(4) = 0
            .TextMatrix(0, 4) = "Estado"
            
            .Col = 5
            .CellAlignment = 4
            .ColWidth(5) = 3400
            .TextMatrix(0, 5) = "Clasificación"
            
            .Col = 6
            .CellAlignment = 4
            .ColWidth(6) = 800
            .TextMatrix(0, 6) = "Incident"
            
            .Col = 7
            .CellAlignment = 4
            .ColWidth(7) = 1000
            .TextMatrix(0, 7) = "CodCliente"
            
            .Col = 8
            .CellAlignment = 4
            .ColWidth(8) = 2400
            .TextMatrix(0, 8) = "Cliente"
            
            .Rows = 1
        End With
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Método de consulta de números
'*****************************************************************************************************R
Public Sub SubConsultarNumeros()
    Dim icompanyid As String
On Error GoTo ErrManager
    ' Si no tenemos código de cliente, asignamos cero
    If Len(Trim(txtiCompanyId.Text)) = 0 Then
        icompanyid = 0
    Else    ' de lo contrario, validamos que sea numérico
        icompanyid = Trim(txtiCompanyId.Text)
        ' si no lo es, mostramos error y salimos
        If Not IsNumeric(icompanyid) Then
            MsgBox "El Código del cliente debe ser numérico.", vbExclamation
            Exit Sub
        End If
    End If
    ' ajustamos la conexion y los parámetros de consulta
    Set Me.VarColNumero.proConexion = Me.proConexion
    ' pendiente de aprobación
    Me.VarColNumero.proEstado = "P"
    ' sin número inicial
    Me.VarColNumero.proNumeroInicial = ""
    ' sin número final
    Me.VarColNumero.proNumeroFinal = ""
    ' asignamos el usuario
    Me.VarColNumero.proUsuario = Me.proUserOnyx
    ' consultamos los números pendientes de aprobación y pintamos la grilla
    If Me.VarColNumero.MetConsultarNumerosSinAprobacion(icompanyid) Then
        Set VarColTempNumeros = New colTemporalAprobacionNumeros
        Call SubFPintarGrid
    Else    ' si ocurre un error, mostramos un mensaje y salimos
        MsgBox "Error al Consultar los números pendientes por aprobar.", vbCritical, App.Title
        Exit Sub
    End If
    ' Si no hay resultados, deshabilitamos los botones
    cmdAprobar.Enabled = (VarColNumero.Count > 0)
    cmdRechazar.Enabled = (VarColNumero.Count > 0)
   Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Método de pintar la grilla
'******************************************************************************************************
Sub SubFPintarGrid()
    Dim Contador  As Integer
    Dim varCuentaColumna As Integer
On Error GoTo ErrManager
    
    Me.grdPendientes.Rows = 1
    'Inicializar Grid
    Call SubFInicializarGrid
   
    Contador = 1
    While Contador <= Me.VarColNumero.Count
        Me.grdPendientes.AddItem Me.VarColNumero.Item(Contador).proRegionCodeDescripcion & vbTab & _
                                    Me.VarColNumero.Item(Contador).proRegionCode & vbTab & _
                                    Me.VarColNumero.Item(Contador).proNumero & vbTab & _
                                    Me.VarColNumero.Item(Contador).proEstadoNumero & vbTab & _
                                    Me.VarColNumero.Item(Contador).proEstadoNumeroDescripcion & vbTab & _
                                    Me.VarColNumero.Item(Contador).proClasificacionDescripcion & vbTab & _
                                    Me.VarColNumero.Item(Contador).proIncidentId & vbTab & _
                                    Me.VarColNumero.Item(Contador).proCompanyId & vbTab & _
                                    Me.VarColNumero.Item(Contador).proCompanyName
        Contador = Contador + 1
    Wend
    Me.grdPendientes.Row = 0
    
    Me.grdProcesados.Rows = 1
    
    Set VarColTempNumeros.varConexion = Me.proConexion
    If Me.VarColTempNumeros.MetConsultarTemporalNumeros(Me.proUserOnyx) Then
    Else
        MsgBox "Error al Consultar los números Procesados.", vbCritical, App.Title
        Exit Sub

    End If
    
    Contador = 1
    While Contador <= Me.VarColTempNumeros.Count
        Me.grdProcesados.AddItem Me.VarColTempNumeros.Item(Contador).temRegionname & vbTab & _
                                    Me.VarColTempNumeros.Item(Contador).temRegionCode & vbTab & _
                                    Me.VarColTempNumeros.Item(Contador).temNumero & vbTab & _
                                    Me.VarColTempNumeros.Item(Contador).temEstadonumero & vbTab & _
                                    Me.VarColTempNumeros.Item(Contador).temDescripcionestado & vbTab & _
                                    Me.VarColTempNumeros.Item(Contador).temClasificacion & vbTab & _
                                    Me.VarColTempNumeros.Item(Contador).temIncidentid & vbTab & _
                                    Me.VarColTempNumeros.Item(Contador).temCompanyid & vbTab & _
                                    Me.VarColTempNumeros.Item(Contador).temCompanyname
        Contador = Contador + 1
    Wend
    'Colorear la Grilla de Procesados
    Contador = 1
    While Contador <= Me.VarColTempNumeros.Count
        Me.grdProcesados.Col = 3
        Me.grdProcesados.Row = Contador
        If Me.grdProcesados.Col = 3 Then
            ' Cambiamos el estado a reservado
            If Me.grdProcesados = "R" Then
            ' Cambiamos el color Aprobado
                For varCuentaColumna = 0 To Me.grdProcesados.Cols - 1
                    Me.grdProcesados.Col = varCuentaColumna
                    Me.grdProcesados.CellBackColor = Me.lbAprobado(5).BackColor
                Next varCuentaColumna
            End If
            If Me.grdProcesados = "L" Then
            ' Cambiamos el color Rechazado
                For varCuentaColumna = 0 To Me.grdProcesados.Cols - 1
                    Me.grdProcesados.Col = varCuentaColumna
                    Me.grdProcesados.CellBackColor = Me.lbRechazado(4).BackColor
                Next varCuentaColumna
            End If
        End If
        Contador = Contador + 1
    Wend

'    Me.grdProcesados.Row = 0
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Método para aprobar números
'******************************************************************************************************
Private Sub Aprobar()
    Dim strEstado As String
On Error GoTo ErrManager
    Me.grdPendientes.Col = 3 ' Columna de estado
    Set Me.VarClasNumero = New EDCAdminVoz.claNumero
    ' Ajustamos la conexión
    Set Me.VarClasNumero.proConexion = Me.proConexion
    strEstado = Me.VarColNumero.Item(Me.grdPendientes.Row).proEstadoNumero
    If Me.grdPendientes.Col = 3 Then
        ' Cambiamos el estado a reservado
        Me.grdPendientes = "R"
        ' Cambiamos el color al de aprobado
        Me.grdPendientes.CellBackColor = Me.lbAprobado(5).BackColor
    End If
    'Pasamos el Número al estado Procesados
    If Me.VarColTempNumeros.MetInsertarTemporalNumeros(Me.proUserOnyx _
        , Me.VarColNumero.Item(Me.grdPendientes.Row).proRegionCode _
        , Me.VarColNumero.Item(Me.grdPendientes.Row).proRegionCodeDescripcion _
        , Me.VarColNumero.Item(Me.grdPendientes.Row).proNumero _
        , "R" _
        , Me.VarColNumero.Item(Me.grdPendientes.Row).proEstadoNumeroDescripcion _
        , Me.VarColNumero.Item(Me.grdPendientes.Row).proClasificacionDescripcion _
        , Me.VarColNumero.Item(Me.grdPendientes.Row).proUpdateBy _
        , Me.proUserOnyx _
        , Me.VarColNumero.Item(Me.grdPendientes.Row).proIncidentId _
        , Me.VarColNumero.Item(Me.grdPendientes.Row).proCompanyId _
        , Me.VarColNumero.Item(Me.grdPendientes.Row).proCompanyName) Then
            Set VarColTempNumeros = New colTemporalAprobacionNumeros
            Call SubFPintarGrid
    Else
        MsgBox "Error al Aprobar el número.", vbCritical, App.Title
        Exit Sub
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


'******************************************************************************************************
'   Método de rechazo masivo de números
'******************************************************************************************************
Private Sub Rechazar()
    Dim strEstado As String
    Dim Contador As Integer
On Error GoTo ErrManager

    Screen.MousePointer = vbHourglass
    Set Me.VarClasNumero = New EDCAdminVoz.claNumero
    Set Me.VarClasNumero.proConexion = Me.proConexion

    Set Me.VarClaNovedadNumero = New claNovedadNumero
    Set Me.VarClaNovedadNumero.proConexion = Me.proConexion

    'Recorre la Coleccion pasa saber cuales han sifdo marcados
    Contador = 1
    While Contador <= Me.VarColNumero.Count
        If Me.VarColNumero.Item(Contador).proSeleccionado = "S" Then
            If Me.grdPendientes.Col = 3 Then
                ' Cambiamos el estado a reservado
                Me.grdPendientes = "L"
                ' Cambiamos el color al de aprobado
                Me.grdPendientes.CellBackColor = Me.lbAprobado(5).BackColor
            End If
        
            If Me.VarColTempNumeros.MetInsertarTemporalNumeros(Me.proUserOnyx _
                , Me.VarColNumero.Item(Me.grdPendientes.Row).proRegionCode _
                , Me.VarColNumero.Item(Me.grdPendientes.Row).proRegionCodeDescripcion _
                , Me.VarColNumero.Item(Me.grdPendientes.Row).proNumero _
                , "L" _
                , Me.VarColNumero.Item(Me.grdPendientes.Row).proEstadoNumeroDescripcion _
                , Me.VarColNumero.Item(Me.grdPendientes.Row).proClasificacionDescripcion _
                , Me.VarColNumero.Item(Me.grdPendientes.Row).proUpdateBy _
                , Me.proUserOnyx _
                , Me.VarColNumero.Item(Me.grdPendientes.Row).proIncidentId _
                , Me.VarColNumero.Item(Me.grdPendientes.Row).proCompanyId _
                , Me.VarColNumero.Item(Me.grdPendientes.Row).proCompanyName) Then
                    Set VarColTempNumeros = New colTemporalAprobacionNumeros
                    Call SubFPintarGrid
            Else
                MsgBox "Error al Rechazar el número.", vbCritical, App.Title
                Screen.MousePointer = vbDefault
                Exit Sub
            End If


'            'Elimina el registro de la Tabla de CT_NOVEDADNUMEROS
'            Me.VarClaNovedadNumero.proRegionCode = Me.VarColNumero.Item(Contador).proRegionCode
'            Me.VarClaNovedadNumero.proNumero = Me.VarColNumero.Item(Contador).proNumero
'            Me.VarClaNovedadNumero.FunGEliminarxNumero
'            ' Liberamos el número
'            Me.VarClasNumero.proEstadoNumero = "L"
'            ' Cambiamos el color al de aprobado
'            Me.grdPendientes.CellBackColor = Me.lbRechazado(4).BackColor
'            Me.VarClasNumero.proUserIdAprobador = Me.proUserOnyx
'            Me.VarClasNumero.proRegionCode = Me.VarColNumero.Item(Contador).proRegionCode
'            Me.VarClasNumero.proNumero = Me.VarColNumero.Item(Contador).proNumero
        'Pasamos el Número al estado Procesados
'            If Not Me.VarClasNumero.FunGModificarEstadoAprobado Then
'                ' Si ocurre un error, lo mostramos
'                MsgBox "Error al actualizar el estado del número.", vbCritical, App.Title
'            Else
'
'                VarColNumerosProcesados.Add Me.proConexion, Me.VarClasNumero.proRecordStatus, _
'                    Me.VarClasNumero.proUpdateDate, Me.VarClasNumero.proUpdateBy, _
'                    Me.VarClasNumero.proClasificacionId, Me.VarClasNumero.proClasificacionDescripcion, _
'                    Me.VarClasNumero.proEstadoNumeroDescripcion, Me.VarClasNumero.proEstadoNumero, _
'                    Me.VarClasNumero.proNumero, Me.VarClasNumero.proRegionCodeDescripcion, _
'                    Me.VarClasNumero.proRegionCode, Me.VarClasNumero.proUserIdAprobador, _
'                    Me.VarClasNumero.proFechaAprobacion, Me.VarClasNumero.proIncidentId, _
'                    Me.VarClasNumero.proCompanyId, Me.VarClasNumero.proCompanyName
'            End If
        End If
        Contador = Contador + 1
    Wend

     'Refresca el grid
    'Call SubConsultarNumeros
    Screen.MousePointer = vbDefault
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Método de limpiar la selección
'******************************************************************************************************
Private Sub subFLimpiarSeleccion()
    Dim varCuenta As Integer
    Dim varCuentaColumna As Integer
On Error GoTo ErrorManager

    For varCuenta = 1 To Me.VarColNumero.Count
        If Me.VarColNumero.Item(varCuenta).proSeleccionado = "S" Then
                Me.VarColNumero.Item(varCuenta).proSeleccionado = False
                Me.grdPendientes.Row = varCuenta
                For varCuentaColumna = 0 To Me.grdPendientes.Cols - 1
                    Me.grdPendientes.Col = varCuentaColumna
                    Me.grdPendientes.CellBackColor = Me.lblNormal.BackColor
                Next varCuentaColumna
        End If
    Next varCuenta
    Exit Sub

ErrorManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Método de cambiar el estado a múltiples números
'******************************************************************************************************
Private Sub CambiarEstadoMasivo()
    Dim strEstado As String
    Dim Contador As Integer
On Error GoTo ErrManager

   If Me.grdPendientes.Rows = 1 Then
        MsgBox "No existen números para ser aprobados.", vbInformation, App.Title
        Exit Sub
    End If

    If Me.grdPendientes.Row = 0 Then
        MsgBox "Debe seleccionar el número que desea aprobar.", vbInformation, App.Title
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Set Me.VarClasNumero = New EDCAdminVoz.claNumero
    Set Me.VarClasNumero.proConexion = Me.proConexion

    'Recorre la Coleccion pasa saber cuales han sifdo marcados
    Contador = 1
    While Contador <= Me.VarColNumero.Count
        If Me.VarColNumero.Item(Contador).proSeleccionado = "S" Then
            'Envia a guardar el estado del número
            Me.VarClasNumero.proEstadoNumero = "R"
            Me.VarClasNumero.proUserIdAprobador = Me.proUserOnyx
            Me.VarClasNumero.proRegionCode = Me.VarColNumero.Item(Contador).proRegionCode
            Me.VarClasNumero.proNumero = Me.VarColNumero.Item(Contador).proNumero
            Me.VarClasNumero.FunGModificarEstadoAprobado
        End If
        Contador = Contador + 1
    Wend

     'Refresca el grid
    Call SubConsultarNumeros
    Screen.MousePointer = vbDefault
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Método para Deshacer números
'******************************************************************************************************
Private Sub Deshacer()
    Dim strEstado As String
    Dim strNumero As String
    Dim Contador As Integer
On Error GoTo ErrManager
    
    Screen.MousePointer = vbHourglass

    If Me.VarColTempNumeros.MetBorrarTemporalNumeros(Me.proUserOnyx _
        , Me.VarColTempNumeros.Item(Me.grdProcesados.Row).temNumero) Then
            Set VarColTempNumeros = New colTemporalAprobacionNumeros
            Call SubFPintarGrid
    Else
        MsgBox "Error al Rechazar el número.", vbCritical, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

'******************************************************************************************************
'   Método para Aprobar/Rechazar Definitivamente los números
'******************************************************************************************************
Private Sub AprobacionFinal(Operacion As String)
    Dim strEstado As String
    Dim strNumero As String
    Dim Contador As Integer
On Error GoTo ErrManager
    
    Screen.MousePointer = vbHourglass

    If Me.VarColTempNumeros.MetDefinitivaTemporalNumeros(Operacion, Me.proUserOnyx) Then
        Screen.MousePointer = vbDefault
    Else
        MsgBox "Error al Aprobar/Rechazar el(los) número(s) Definitivamente", vbCritical, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub mnuAprobar_Click()
On Error GoTo ErrManager
    Call Aprobar
        Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub mnuRechazar_Click()
On Error GoTo ErrManager
    Call Rechazar
        Exit Sub
ErrManager:
    SubGMuestraError

End Sub


