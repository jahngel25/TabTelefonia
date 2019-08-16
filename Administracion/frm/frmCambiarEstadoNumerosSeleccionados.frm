VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmCambiarEstadoNumerosSeleccionados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Estados"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8955
   Icon            =   "frmCambiarEstadoNumerosSeleccionados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   7080
      TabIndex        =   5
      Top             =   1200
      Width           =   1500
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   2160
      Left            =   0
      TabIndex        =   0
      Top             =   255
      Width           =   8955
      _Version        =   65536
      _ExtentX        =   15796
      _ExtentY        =   3810
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
      BorderWidth     =   2
      BevelOuter      =   1
      Begin VB.Frame Frame1 
         Height          =   1605
         Left            =   3720
         TabIndex        =   1
         Top             =   -75
         Width           =   5205
         Begin VB.CommandButton cmdCambiarEstado 
            Caption         =   "&Cambiar"
            Height          =   315
            Left            =   3360
            TabIndex        =   4
            Top             =   540
            Width           =   1500
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   255
            Left            =   15
            TabIndex        =   2
            Top             =   90
            Width           =   5145
            _Version        =   65536
            _ExtentX        =   9075
            _ExtentY        =   450
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   12620376
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
      End
      Begin VB.Frame Frame2 
         Height          =   2025
         Left            =   45
         TabIndex        =   12
         Top             =   -15
         Width           =   3570
         Begin VB.ComboBox cboNombreEstado 
            Height          =   315
            Left            =   195
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   570
            Width           =   2685
         End
         Begin VB.ComboBox cboCodigoEstado 
            Height          =   315
            Left            =   1455
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   585
            Visible         =   0   'False
            Width           =   765
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   30
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   450
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   12620376
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
         Begin VB.Label lblEstado 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   210
            TabIndex        =   15
            Top             =   330
            Width           =   540
         End
      End
      Begin VB.Label lblMensaje 
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Debe seleccionar un estado y luego dar click en cambiar para realizar el respectivo cambio en los numeros de la grilla."
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   3720
         TabIndex        =   6
         Top             =   1545
         Width           =   5250
      End
   End
   Begin Threed.SSPanel pnlTitulo 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8955
      _Version        =   65536
      _ExtentX        =   15796
      _ExtentY        =   450
      _StockProps     =   15
      ForeColor       =   16777215
      BackColor       =   12620376
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
   Begin Threed.SSPanel SSPanel5 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   8925
      _Version        =   65536
      _ExtentX        =   15743
      _ExtentY        =   450
      _StockProps     =   15
      Caption         =   "Numeros a cambiar"
      ForeColor       =   16777215
      BackColor       =   12620376
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
   Begin MSFlexGridLib.MSFlexGrid grdNumeros 
      Height          =   4920
      Left            =   -15
      TabIndex        =   8
      Top             =   2685
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   8678
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   7575
      Width           =   9015
      _Version        =   65536
      _ExtentX        =   15901
      _ExtentY        =   741
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
      Begin VB.TextBox txtCantidadSeleccionados 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   10
         Top             =   60
         Width           =   705
      End
      Begin VB.Label lblCantidadRegistrosSeleccion 
         BackColor       =   &H00C8D0D4&
         Caption         =   "Cantidad Registros"
         Height          =   195
         Left            =   5640
         TabIndex        =   11
         Top             =   120
         Width           =   2445
      End
   End
End
Attribute VB_Name = "frmCambiarEstadoNumerosSeleccionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'       DESCRIPCION: Formulario que permite cambiar el estado de los
'       numeros seleccionados en la pantalla de consulta de numeros
'       Autor: TOPGROUP S.A.
'       Fecha: 10/10/2009
'       Version:              1.0.000
'       Requerimiento:        5322
'*******************************************************************
Option Explicit

Public proNumeros As colNumero


'Propiedad de conexion
Public proConexion As ADODB.Connection
Public proLogin As String
Public proGuardado As Boolean

Private varEstadosDestinos As colEstadoOrigenDestino
Private varCantidadNumerosACambiar As Integer
Private varPintarObservaciones As Boolean
Private varMensajeResultado
Private Const ConstEstadoOrigen = "Origen"
Private Const ConstEstadoDestino = "Destino"

Private Sub SubFInicializarGridNumeros()
    On Error GoTo ErrManager:
    
    With Me.grdNumeros
        .Cols = 10
        .Rows = 1
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 0
        .TextMatrix(0, 0) = "Codigo Ciudad"
        
        .Col = 1
        .CellAlignment = 4
        .ColWidth(1) = 1425
        .TextMatrix(0, 1) = "Ciudad"
        
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 1215
        .TextMatrix(0, 2) = "Numero"
        
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 0
        .TextMatrix(0, 3) = "Codigo Estado"
        
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 960
        .TextMatrix(0, 4) = "Estado"
        
        .Col = 5
        .CellAlignment = 4
        .ColWidth(5) = 0
        .TextMatrix(0, 5) = "Codigo Clasificacion"
        
        .Col = 6
        .CellAlignment = 4
        .ColWidth(6) = 1800
        .TextMatrix(0, 6) = "Clasificacion"
        
        .Col = 7
        .CellAlignment = 4
        .ColWidth(7) = 1455
        .TextMatrix(0, 7) = "Usuario"
        
        .Col = 8
        .CellAlignment = 4
        .ColWidth(8) = 2040
        .TextMatrix(0, 8) = "Fecha"
        
        .Col = 9
        .CellAlignment = 4
        .ColWidth(9) = 0
        .TextMatrix(0, 9) = "Observaciones"
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Sub SubFLlenarComboEstados()
On Error GoTo ErrorManager
    Dim i As Integer
    
    Me.cboNombreEstado.Clear
    Me.cboCodigoEstado.Clear

    For i = 1 To cboNombreEstado.ListCount
        cboNombreEstado.RemoveItem i
        cboCodigoEstado.RemoveItem i
    Next
    
    For i = 1 To varEstadosDestinos.Count
        cboNombreEstado.AddItem varEstadosDestinos.Item(i).proDescripcion
        cboCodigoEstado.AddItem varEstadosDestinos.Item(i).proEstadoNumero
    Next
    
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub
Private Sub SubFPintarGridNumeros()
    Dim varContador As Integer
    Dim varContadorAux As Integer
    On Error GoTo ErrManager
    
    If varPintarObservaciones = True Then
        Me.grdNumeros.ColWidth(9) = 4000
    Else
        Me.grdNumeros.ColWidth(9) = 0
    End If
    varCantidadNumerosACambiar = 0
    Me.grdNumeros.Rows = 1
    Me.grdNumeros.Redraw = False
    For varContador = 1 To Me.proNumeros.Count
        If Me.proNumeros.Item(varContador).proSeleccionado = "S" Then
                Dim varTmp As String
                varTmp = Me.proNumeros.Item(varContador).proRegionCode & vbTab & _
                           Me.proNumeros.Item(varContador).proRegionCodeDescripcion & vbTab & _
                           Me.proNumeros.Item(varContador).proNumero & vbTab
                If varPintarObservaciones = True Then
                    varTmp = varTmp & Me.proNumeros.Item(varContador).proEstadoNumero & vbTab & _
                    Me.proNumeros.Item(varContador).proEstadoNumeroDescripcion
                    
                    varTmp = varTmp & vbTab & Me.proNumeros.Item(varContador).proClasificacionId & vbTab & _
                           Me.proNumeros.Item(varContador).proClasificacionDescripcion & vbTab & _
                           Me.proNumeros.Item(varContador).proUpdateBy & vbTab & _
                           Me.proNumeros.Item(varContador).proUpdateDate
                    varTmp = varTmp & vbTab & Me.proNumeros.Item(varContador).proObservacionesCambioEstado
                    If Left(Me.proNumeros.Item(varContador).proObservacionesCambioEstado, 5) <> "Error" Then
                           'varTmp = varTmp & Me.cboCodigoEstado.Text & vbTab & _
                           'Me.cboNombreEstado.Text
                           Me.proGuardado = True
                    End If
                Else
                    varTmp = varTmp & Me.proNumeros.Item(varContador).proEstadoNumero & vbTab & _
                           Me.proNumeros.Item(varContador).proEstadoNumeroDescripcion
                    varTmp = varTmp & vbTab & Me.proNumeros.Item(varContador).proClasificacionId & vbTab & _
                           Me.proNumeros.Item(varContador).proClasificacionDescripcion & vbTab & _
                           Me.proNumeros.Item(varContador).proUpdateBy & vbTab & _
                           Me.proNumeros.Item(varContador).proUpdateDate
                    varTmp = varTmp & vbTab & "--"
                End If


                Me.grdNumeros.AddItem varTmp
                Me.grdNumeros.CellBackColor = &HC0FFFF
                varCantidadNumerosACambiar = varCantidadNumerosACambiar + 1
        End If
                              
'        If Me.proNumeros.Item(varContador).proSeleccionado = "S" Then
'            Me.grdNumeros.Row = Me.grdNumeros.Rows - 1
'            For varContadorAux = 0 To Me.grdNumeros.Cols - 1
'                Me.grdNumeros.Col = varContadorAux
'                Me.grdNumeros.CellBackColor = &HC0FFFF
'            Next varContadorAux
'        End If
                              
    Next varContador
    
    Me.grdNumeros.Row = 0
    Me.grdNumeros.Col = 0
    Me.grdNumeros.Redraw = True
    Me.txtCantidadSeleccionados.Text = varCantidadNumerosACambiar

    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cboNombreEstado_Click()
    Me.cboCodigoEstado.ListIndex = Me.cboNombreEstado.ListIndex
End Sub


Private Sub cmdCambiarEstado_Click()

    Dim mensajeResultado As String
    
    ' primero validar que haya un estado seleccionado
    If Me.cboNombreEstado.Text = "" Then
        MsgBox "Debe seleccionar un estado para poder realizar el cambio.", vbInformation + vbOKOnly, "Cambio de Estado"
        Exit Sub
    End If
    
    ' despues mostrar mensaje de confirmacion
    If (varCantidadNumerosACambiar < 1) Then
        MsgBox "No existen numeros seleccionados.", vbCritical, "Problema"
        Exit Sub
    Else
        If MsgBox("Se dispone a cambiar el estado de " & varCantidadNumerosACambiar & " números, desea continuar?", vbYesNo + vbQuestion, "Atención") = vbNo Then
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' despues invocar al metodo que hace el cambio de estado
    mensajeResultado = proNumeros.MetCambiarEstadoNumeros(Me.cboCodigoEstado.Text, Me.cboNombreEstado.Text, Me.proLogin)
    
    ' despues repintar la grilla pero mostrando las observaciones
    varPintarObservaciones = True
    'If Me.proNumeros.MetConsultarNumeros Then
        Call SubFPintarGridNumeros
    'Else
    '    MsgBox "Error al consultar los números.", vbCritical, App.Title
    '    Screen.MousePointer = 0
    '    Exit Sub
    'End If
    Dim anchoAGanar As Integer
    anchoAGanar = 4000
    Me.Width = Me.Width + anchoAGanar
    Me.pnlTitulo.Width = Me.pnlTitulo.Width + anchoAGanar
    Me.Frame1.Width = Me.Frame1.Width + anchoAGanar
    Me.SSPanel3.Width = Me.SSPanel3.Width + anchoAGanar
    Me.lblMensaje.Width = Me.lblMensaje.Width + anchoAGanar
    Me.SSPanel1.Width = Me.SSPanel1.Width + anchoAGanar
    Me.SSPanel5.Width = Me.SSPanel5.Width + anchoAGanar
    Me.grdNumeros.Width = Me.grdNumeros.Width + anchoAGanar
    Me.SSPanel2.Width = Me.SSPanel2.Width + anchoAGanar

    Screen.MousePointer = vbDefault
    
    ' despues mostrar un mensaje con el status de la operacion
    MsgBox mensajeResultado, vbOKOnly + vbInformation, "Cambio de Estados"

End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrorManager
        
    'Inicializar combo de Estados Destinos
    Set varEstadosDestinos = New colEstadoOrigenDestino
    Set varEstadosDestinos.proConexion = Me.proConexion
    varCantidadNumerosACambiar = 0
    varPintarObservaciones = False
    varEstadosDestinos.proTipoEstado = ConstEstadoDestino
    Me.proGuardado = False
    
    If varEstadosDestinos.FunGConsultaEstadosPorTipo Then
        Call SubFLlenarComboEstados
    Else
        MsgBox "Error al consultar los estados.", vbCritical, App.Title
        Exit Sub
    End If
    
    'Cargar grilla con los numeros seleccionados en la pantalla anterior
    Call SubFInicializarGridNumeros

    Call SubFPintarGridNumeros

    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

