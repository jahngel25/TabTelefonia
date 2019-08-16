VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmValorServicioSup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Valor Servicio Suplementario"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   315
      Left            =   3600
      TabIndex        =   9
      Top             =   1680
      Width           =   1545
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   1365
   End
   Begin Threed.SSPanel SSPanelCombo 
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   5145
      _Version        =   65536
      _ExtentX        =   9075
      _ExtentY        =   1244
      _StockProps     =   15
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.ComboBox cboNombre 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   4125
      End
      Begin VB.ComboBox cboCodigo 
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblColumna 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   240
         Width           =   405
      End
   End
   Begin Threed.SSPanel SSPanelTexto 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   5145
      _Version        =   65536
      _ExtentX        =   9075
      _ExtentY        =   1296
      _StockProps     =   15
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtValor 
         Height          =   435
         Left            =   930
         TabIndex        =   7
         Top             =   120
         Width           =   4065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   270
         TabIndex        =   6
         Top             =   240
         Width           =   405
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   270
      Width           =   5145
      _Version        =   65536
      _ExtentX        =   9075
      _ExtentY        =   1296
      _StockProps     =   15
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmValorServicioSup.frx":0000
         Top             =   60
         Width           =   480
      End
      Begin VB.Label lblServicioSuplementario 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   630
         TabIndex        =   15
         Top             =   60
         Width           =   4425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C09258&
         BackStyle       =   0  'Transparent
         Caption         =   "Región : "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2910
         TabIndex        =   14
         Top             =   330
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C09258&
         BackStyle       =   0  'Transparent
         Caption         =   "Línea : "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   630
         TabIndex        =   13
         Top             =   330
         Width           =   555
      End
      Begin VB.Label lblRegion 
         AutoSize        =   -1  'True
         BackColor       =   &H00C09258&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3570
         TabIndex        =   12
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label lblLinea 
         AutoSize        =   -1  'True
         BackColor       =   &H00C09258&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1170
         TabIndex        =   11
         Top             =   330
         Width           =   1605
      End
   End
   Begin VB.Frame fraTituloProducto 
      BackColor       =   &H00C09258&
      Caption         =   "  Información  de la línea"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   -30
      TabIndex        =   0
      Top             =   30
      Width           =   5175
   End
End
Attribute VB_Name = "frmValorServicioSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proConexion As ADODB.Connection
Public proServicioSuplementario As String
Public proNombreServicioSuplementario As String
Public proNovedadId  As String
Public proTipoServicio As String
Public proTelefono As String
Public proRegion As String
Public prEsNovedad As String 'Para saber si se almacena la info en CT_NOVEDADVALORSERVICIOXNUMERO o en la CT_VALORSERVICIOXNUMERO
Public proValor As String
Public proTodos As Integer


Public varColValorServicio As EDCAdminVoz.colValorServicio
Public varClaValorServicioxNumero As claValorServicioxnumero
Public varClaNovedadvalorServicioxNumero As claNovedadValorServicioxNumero

Private Sub cmdEliminar_Click()
On Error GoTo ErrManager

    If proTodos = 1 Then 'El cambio lo deben sufrir todos los numeros
        If MsgBox("Desea eliminar todos los valores de este servicio suplementario?", vbYesNo + vbInformation, App.Title) = vbNo Then
            Exit Sub
        End If
    End If

    If Me.proNovedadId = 0 Then
        'Quiere decir que debo eliminar la información en la tabla CT_VALORSERVICIOXNUMERO
        Set varClaValorServicioxNumero = New claValorServicioxnumero
        Set varClaValorServicioxNumero.proConexion = Me.proConexion
        varClaValorServicioxNumero.proServicioSuplementario = Me.proServicioSuplementario
        varClaValorServicioxNumero.proNumero = Me.proTelefono
        varClaValorServicioxNumero.proRegionCode = Me.proRegion
        If proTodos = 1 Then 'El cambio lo deben sufrir todos los numeros
            If varClaValorServicioxNumero.FunGEliminarTodos Then
                MsgBox "Se elimino el valor para el servicio exitosamente.", vbInformation, App.Title
                frmEdicionServicios.proAccionValor = 2
                Unload Me
            Else
                MsgBox "No fue posible eliminar el valor para el servicio suplementario.", vbInformation, App.Title
            End If
        Else
            If varClaValorServicioxNumero.FunGEliminar Then
                MsgBox "Se elimino el valor para el servicio exitosamente.", vbInformation, App.Title
                frmEdicionServicios.proAccionValor = 2
                Unload Me
            Else
                MsgBox "No fue posible eliminar el valor para el servicio suplementario.", vbInformation, App.Title
            End If
        End If
    Else
        'Quiere decir que debo eliminar la información en la tabla CT_NOVEDADVALORSERVICIOXNUMERO
        Set varClaNovedadvalorServicioxNumero = New claNovedadValorServicioxNumero
        Set varClaNovedadvalorServicioxNumero.proConexion = Me.proConexion
        varClaNovedadvalorServicioxNumero.proNovedadNumeroId = Me.proNovedadId
        varClaNovedadvalorServicioxNumero.proServicioSuplementario = Me.proServicioSuplementario
        varClaNovedadvalorServicioxNumero.proNumero = Me.proTelefono
        varClaNovedadvalorServicioxNumero.proRegion = Me.proRegion
        If proTodos = 1 Then 'El cambio lo deben sufrir todos los numeros
            If varClaNovedadvalorServicioxNumero.FunGEliminarTodos Then
                MsgBox "Se elimino el valor para el servicio exitosamente.", vbInformation, App.Title
                frmEdicionServicios.proAccionValor = 2
                Unload Me
            Else
                MsgBox "No fue posible eliminar el valor para el servicio suplementario.", vbInformation, App.Title
            End If
        Else
            If varClaNovedadvalorServicioxNumero.FunGEliminar Then
                MsgBox "Se elimino el valor para el servicio exitosamente.", vbInformation, App.Title
                frmEdicionServicios.proAccionValor = 2
                Unload Me
            Else
                MsgBox "No fue posible eliminar el valor para el servicio suplementario.", vbInformation, App.Title
            End If
        End If

    End If
    Exit Sub
ErrManager:
    SubGMuestraError

End Sub

Private Sub cmdGuardar_Click()
On Error GoTo ErrManager
    If proTodos = 1 Then 'El cambio lo deben sufrir todos los numeros
        If MsgBox("Desea aplicar los cambios a todos los valores de este servicio suplementario?", vbYesNo + vbInformation, App.Title) = vbNo Then
            Exit Sub
        End If
    End If


    If Me.proNovedadId = 0 Then
        'Quiere decir que debo almacenar la información en la tabla CT_VALORSERVICIOXNUMERO
        Set varClaValorServicioxNumero = New claValorServicioxnumero
        Set varClaValorServicioxNumero.proConexion = Me.proConexion
        varClaValorServicioxNumero.proServicioSuplementario = Me.proServicioSuplementario
        varClaValorServicioxNumero.proNumero = Me.proTelefono
        varClaValorServicioxNumero.proRegionCode = Me.proRegion
        If proTipoServicio = "L" Then
            If cboNombre.ListIndex = -1 Then
                MsgBox "Debe seleccionar un valor.", vbInformation, App.Title
                Exit Sub
            Else
                varClaValorServicioxNumero.proValor = Me.cboNombre.List(cboNombre.ListIndex)
            End If
        Else
            If Trim(Me.txtValor.Text) <> "" Then
                varClaValorServicioxNumero.proValor = Trim(Me.txtValor.Text)
            Else
                MsgBox "Debe ingresar un valor.", vbInformation, App.Title
                Me.txtValor.SetFocus
                Exit Sub
            End If
        End If
        If IsNull(proValor) = False And proValor <> "" And proValor <> "p" And proValor <> "q" Then
            If proTodos = 1 Then 'El cambio lo deben sufrir todos los numeros
                If varClaValorServicioxNumero.FunGInsertarTodos Then
                    MsgBox "Se almaceno el valor para el servicio exitosamente.", vbInformation, App.Title
                    frmEdicionServicios.proAccionValor = 1
                    Unload Me
                Else
                    MsgBox "No fue posible almacenar el valor para el servicio suplementario.", vbInformation, App.Title
                End If
            Else
                If varClaValorServicioxNumero.FunGModificar Then
                    MsgBox "Se almaceno el valor para el servicio exitosamente.", vbInformation, App.Title
                    frmEdicionServicios.proAccionValor = 1
                    Unload Me
                Else
                    MsgBox "No fue posible almacenar el valor para el servicio suplementario.", vbInformation, App.Title
                End If
            End If
                
        Else
            If proTodos = 1 Then 'El cambio lo deben sufrir todos los numeros
                If varClaValorServicioxNumero.FunGInsertarTodos Then
                    MsgBox "Se almaceno el valor para el servicio exitosamente.", vbInformation, App.Title
                    frmEdicionServicios.proAccionValor = 1
                    Unload Me
                Else
                    MsgBox "No fue posible almacenar el valor para el servicio suplementario.", vbInformation, App.Title
                End If
            Else
                If varClaValorServicioxNumero.FunGInsertar Then
                    MsgBox "Se almaceno el valor para el servicio exitosamente.", vbInformation, App.Title
                    frmEdicionServicios.proAccionValor = 1
                    Unload Me
                Else
                    MsgBox "No fue posible almacenar el valor para el servicio suplementario.", vbInformation, App.Title
                End If
            End If
        End If
        
    Else
        'Quiere decir que debo almacenar la información en la tabla CT_NOVEDADVALORSERVICIOXNUMERO
        Set varClaNovedadvalorServicioxNumero = New claNovedadValorServicioxNumero
        Set varClaNovedadvalorServicioxNumero.proConexion = Me.proConexion
        varClaNovedadvalorServicioxNumero.proNovedadNumeroId = Me.proNovedadId
        varClaNovedadvalorServicioxNumero.proServicioSuplementario = Me.proServicioSuplementario
        varClaNovedadvalorServicioxNumero.proNumero = Me.proTelefono
        varClaNovedadvalorServicioxNumero.proRegion = Me.proRegion
        If proTipoServicio = "L" Then
            If cboNombre.ListIndex = -1 Then
                MsgBox "Debe seleccionar un valor.", vbInformation, App.Title
                Exit Sub
            Else
                varClaNovedadvalorServicioxNumero.proValor = Me.cboNombre.List(cboNombre.ListIndex)
            End If
        Else
            If Trim(Me.txtValor.Text) <> "" Then
                varClaNovedadvalorServicioxNumero.proValor = Trim(Me.txtValor.Text)
            Else
                MsgBox "Debe ingresar un valor.", vbInformation, App.Title
                Me.txtValor.SetFocus
                Exit Sub
            End If
        End If
        If IsNull(proValor) = False And proValor <> "" And proValor <> "p" And proValor <> "q" Then
            If proTodos = 1 Then 'El cambio lo deben sufrir todos los numeros
                If varClaNovedadvalorServicioxNumero.FunGInsertarTodos Then
                    MsgBox "Se almaceno el valor para el servicio exitosamente.", vbInformation, App.Title
                    frmEdicionServicios.proAccionValor = 1
                    Unload Me
                Else
                    MsgBox "No fue posible almacenar el valor para el servicio suplementario.", vbInformation, App.Title
                End If
            Else
                If varClaNovedadvalorServicioxNumero.FunGModificar Then
                    MsgBox "Se almaceno el valor para el servicio exitosamente.", vbInformation, App.Title
                    frmEdicionServicios.proAccionValor = 1
                    Unload Me
                Else
                    MsgBox "No fue posible almacenar el valor para el servicio suplementario.", vbInformation, App.Title
                End If
            End If
        Else
            If proTodos = 1 Then 'El cambio lo deben sufrir todos los numeros
                If varClaNovedadvalorServicioxNumero.FunGInsertarTodos Then
                    MsgBox "Se almaceno el valor para el servicio exitosamente.", vbInformation, App.Title
                    frmEdicionServicios.proAccionValor = 1
                    Unload Me
                Else
                    MsgBox "No fue posible almacenar el valor para el servicio suplementario.", vbInformation, App.Title
                End If
            Else
                If varClaNovedadvalorServicioxNumero.FunGInsertar Then
                    MsgBox "Se almaceno el valor para el servicio exitosamente.", vbInformation, App.Title
                    frmEdicionServicios.proAccionValor = 1
                    Unload Me
                Else
                    MsgBox "No fue posible almacenar el valor para el servicio suplementario.", vbInformation, App.Title
                End If
            End If
        End If

    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrManager
    If KeyAscii = 27 Then ' scape
        Unload Me
    End If
    If KeyAscii = 13 Then ' enter
        Call cmdGuardar_Click
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
     On Error GoTo ErrManager
    Me.Top = frmEdicionServicios.Top + 600
    Me.Left = frmEdicionServicios.Left + 150
    Me.lblLinea.Caption = Me.proTelefono
    Me.lblRegion.Caption = Me.proRegion
    Me.lblServicioSuplementario.Caption = Me.proNombreServicioSuplementario
    '
    If proTipoServicio = "T" Then
        'Hace visible el panel de texto
        Me.SSPanelTexto.Visible = True
        'Trae el valor si es una modificacion
        If IsNull(proValor) = False And proValor <> "" And proValor <> "p" And proValor <> "q" Then
            Me.txtValor.Text = Me.proValor
        End If
    ElseIf proTipoServicio = "L" Then
        'hace visible el panel de combo
        Me.SSPanelCombo.Visible = True
        'Trae los posibles valores con el valor que trae o con el valor por default
        Set varColValorServicio = New EDCAdminVoz.colValorServicio
        Set varColValorServicio.proConexion = Me.proConexion
        varColValorServicio.proServicioSuplementarioId = Me.proServicioSuplementario
        If varColValorServicio.FunGConsulta Then
            Call SubFPintarComboValores
        Else
            MsgBox "No fue posible consultar los valores para el servicio suplementario.", vbInformation, App.Title
        End If
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub SubFPintarComboValores()
    
    Dim varContador As Integer
    Dim iIndice As Integer
    Dim iDefault As Integer
    On Error GoTo ErrorManager
    
    Me.cboCodigo.Clear
    Me.cboNombre.Clear
    iIndice = 0
    For varContador = 1 To Me.varColValorServicio.Count
        If IsNull(proValor) = False And proValor <> "" And proValor <> "p" And proValor <> "q" Then 'Viene un valor de la otra ventana
            If proValor = Me.varColValorServicio.Item(varContador).proValor Then
                iIndice = varContador
            End If
        End If
        If Me.varColValorServicio.Item(varContador).proDefault = "True" Then
            iDefault = varContador
        End If
        Me.cboNombre.AddItem Me.varColValorServicio.Item(varContador).proValor
        Me.cboCodigo.AddItem Me.varColValorServicio.Item(varContador).proServicioSuplementarioId
    Next varContador
    If iIndice = 0 Then
        Me.cboCodigo.ListIndex = iDefault - 1
        Me.cboNombre.ListIndex = iDefault - 1
    Else
        Me.cboCodigo.ListIndex = iIndice - 1
        Me.cboNombre.ListIndex = iIndice - 1
    End If
   
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub
