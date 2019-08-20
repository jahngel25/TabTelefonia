VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmCamposProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrar campos de un producto"
   ClientHeight    =   3585
   ClientLeft      =   3780
   ClientTop       =   2475
   ClientWidth     =   4980
   Icon            =   "frmCamposProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4980
   Begin VB.Frame fraFondoGrid 
      Height          =   2865
      Left            =   30
      TabIndex        =   20
      Top             =   4830
      Width           =   9945
      Begin MSFlexGridLib.MSFlexGrid grdCampos 
         Height          =   2715
         Left            =   30
         TabIndex        =   3
         Top             =   120
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4789
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame fraBotones 
      Height          =   525
      Left            =   60
      TabIndex        =   21
      Top             =   3030
      Width           =   4875
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   315
         Left            =   2340
         TabIndex        =   6
         Top             =   150
         Width           =   1065
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   315
         Left            =   1230
         TabIndex        =   5
         Top             =   150
         Width           =   1065
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   150
         Width           =   1065
      End
   End
   Begin VB.Frame fraFondoEdicion 
      Height          =   3015
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   10155
      Begin VB.CommandButton cmdValores 
         Caption         =   "&Valores"
         Height          =   285
         Left            =   7710
         TabIndex        =   42
         Top             =   2640
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   8850
         TabIndex        =   15
         Top             =   2640
         Width           =   1005
      End
      Begin Threed.SSPanel pnlCampo 
         Height          =   2865
         Left            =   0
         TabIndex        =   23
         Top             =   90
         Width           =   4905
         _Version        =   65536
         _ExtentX        =   8652
         _ExtentY        =   5054
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
         BevelInner      =   2
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "&Guardar"
            Height          =   315
            Left            =   3810
            TabIndex        =   43
            Top             =   2490
            Width           =   1005
         End
         Begin VB.CheckBox chkValidarRepetidos 
            Alignment       =   1  'Right Justify
            Caption         =   "Validar Repetidos"
            Height          =   285
            Left            =   3240
            TabIndex        =   41
            Top             =   1710
            Width           =   1545
         End
         Begin VB.CheckBox chkObligatorioOT 
            Caption         =   "OT"
            Height          =   195
            Left            =   1020
            TabIndex        =   40
            Top             =   2550
            Width           =   1035
         End
         Begin VB.CheckBox chkObligatorioAtencion 
            Caption         =   "Atención"
            Height          =   195
            Left            =   1020
            TabIndex        =   39
            Top             =   2280
            Width           =   1035
         End
         Begin VB.ComboBox cboCodigoMascara 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1350
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.ComboBox cboCodigoTipo 
            Height          =   315
            ItemData        =   "frmCamposProducto.frx":0CCA
            Left            =   990
            List            =   "frmCamposProducto.frx":0CCC
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1020
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.ComboBox cboMascara 
            Enabled         =   0   'False
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1350
            Width           =   3825
         End
         Begin VB.ComboBox cboCampos 
            Height          =   315
            ItemData        =   "frmCamposProducto.frx":0CCE
            Left            =   990
            List            =   "frmCamposProducto.frx":0CD0
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
            Width           =   3840
         End
         Begin VB.TextBox txtEtiqueta 
            Height          =   315
            Left            =   990
            TabIndex        =   8
            Top             =   690
            Width           =   3825
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1020
            Width           =   3825
         End
         Begin VB.TextBox txtTamano 
            Enabled         =   0   'False
            Height          =   285
            Left            =   990
            TabIndex        =   11
            Top             =   1680
            Width           =   705
         End
         Begin VB.CheckBox chkObligatorioVenta 
            Caption         =   "Venta"
            Height          =   195
            Left            =   1020
            TabIndex        =   12
            Top             =   2010
            Width           =   1035
         End
         Begin Threed.SSPanel pnlTituloCampo 
            Height          =   255
            Left            =   30
            TabIndex        =   34
            Top             =   60
            Width           =   4785
            _Version        =   65536
            _ExtentX        =   8440
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Definición del campo"
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
         Begin VB.Label lblObligatorio 
            AutoSize        =   -1  'True
            Caption         =   "Editable en:"
            Height          =   195
            Left            =   150
            TabIndex        =   29
            Top             =   2010
            Width           =   840
         End
         Begin VB.Label lblMascara 
            AutoSize        =   -1  'True
            Caption         =   "Máscara:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1350
            Width           =   660
         End
         Begin VB.Label lblCampo 
            AutoSize        =   -1  'True
            Caption         =   "Campo:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   390
            Width           =   540
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Etiqueta:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lblTipo 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1050
            Width           =   360
         End
         Begin VB.Label lblTamano 
            Caption         =   "Tamaño:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1680
            Width           =   645
         End
      End
      Begin Threed.SSPanel pnlInterfase 
         Height          =   2415
         Left            =   5040
         TabIndex        =   30
         Top             =   150
         Visible         =   0   'False
         Width           =   5025
         _Version        =   65536
         _ExtentX        =   8864
         _ExtentY        =   4260
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
         BevelInner      =   2
         Begin Threed.SSPanel pnlTituloInterfase 
            Height          =   255
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   4905
            _Version        =   65536
            _ExtentX        =   8652
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Interfase  Billing"
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
         Begin VB.TextBox txtPosicion 
            Enabled         =   0   'False
            Height          =   285
            Left            =   930
            MaxLength       =   2
            TabIndex        =   14
            Top             =   1530
            Width           =   405
         End
         Begin VB.CheckBox chkInterfase 
            Height          =   195
            Left            =   960
            TabIndex        =   13
            Top             =   540
            Width           =   195
         End
         Begin VB.Label lblDescripcionPosicion 
            Caption         =   "Indica la posición en la cual se debe enviar el campo dentro de la interfase."
            Height          =   645
            Left            =   1500
            TabIndex        =   36
            Top             =   1500
            Width           =   3075
         End
         Begin VB.Label lblDescripcionBilling 
            Caption         =   "Si se encuentra habilitado indica que este campo debe se enviado en la interfase con Billing."
            Height          =   645
            Left            =   1500
            TabIndex        =   35
            Top             =   480
            Width           =   3075
         End
         Begin VB.Label lblPosicionInterfase 
            AutoSize        =   -1  'True
            Caption         =   "Posición:"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   1530
            Width           =   645
         End
         Begin VB.Label lblInterfase 
            AutoSize        =   -1  'True
            Caption         =   "Enviar:"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   510
            Width           =   525
         End
      End
   End
   Begin VB.Frame FraFiltro 
      BackColor       =   &H00C09258&
      Caption         =   "  Filtro "
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
      Left            =   30
      TabIndex        =   0
      Top             =   3690
      Width           =   9945
   End
   Begin VB.Frame fraFondoFiltro 
      Height          =   555
      Left            =   30
      TabIndex        =   16
      Top             =   3990
      Width           =   9915
      Begin VB.ComboBox cboCodigoProducto 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   150
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox cboNombreProducto 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   8265
      End
      Begin VB.CommandButton cmBuscarCampos 
         Caption         =   "&Buscar Campos"
         Height          =   285
         Left            =   7650
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblProducto 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   210
         Width           =   690
      End
   End
   Begin VB.Frame fraTituloCampos 
      BackColor       =   &H00C09258&
      Caption         =   "  Campos  "
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
      Left            =   0
      TabIndex        =   19
      Top             =   4560
      Width           =   9975
   End
End
Attribute VB_Name = "frmCamposProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Esta forma se modificó de su forma original para poder ajustarla a la parametrización jerarquica
'los frames que sobraron únicamente se ocultaron
'I&T Germán A.F.G. - 16 sep 2004
Option Explicit

Public proConexion As ADODB.Connection
Public proProducto As colProductMaster
Public proParametrosProducto As colParametroProducto
Public proclaParametrosProducto As claParametroProducto
'Parametros de la forma
Public proProductNumber As String
Public proCampo As String
Public proCampoPadre  As String
Public Sub PriductListIndex(LisIndex As Long)
    On Error GoTo ErrManager
    Me.cboCodigoProducto.ListIndex = LisIndex
    Me.cboNombreProducto.ListIndex = LisIndex
    Call cmdNuevo_Click
    Exit Sub
ErrManager:
    SubGMuestraError

End Sub
Public Sub Insertar(valor As Boolean)
    On Error GoTo ErrManager

    Me.cboCampos.Enabled = valor
    
Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cboMascara_Click()
    On Error GoTo ErrManager
    Me.cboCodigoMascara.ListIndex = Me.cboMascara.ListIndex
   
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cboNombreProducto_Click()
    On Error GoTo ErrManager
    
    Me.cboCodigoProducto.ListIndex = Me.cboNombreProducto.ListIndex
    
    Call cmBuscarCampos_Click
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub cboTipo_Click()
    On Error GoTo ErrManager
        
        Me.cboCodigoTipo.ListIndex = Me.cboTipo.ListIndex
        
        Select Case Me.cboCodigoTipo.Text
            Case "T"
                Me.cboMascara.Enabled = True
                Me.txtTamano.Enabled = True
                Me.cmdValores.Visible = False
                Me.chkValidarRepetidos.Value = 0
                Me.chkValidarRepetidos.Enabled = False
            Case "L"
                Me.cboMascara.ListIndex = -1
                Me.txtTamano.Text = ""
                Me.txtTamano.Enabled = False
                Me.cboMascara.Enabled = False
                Me.cmdValores.Visible = True
                Me.cmdValores.Enabled = False
                Me.chkValidarRepetidos.Enabled = True
            Case "F"
                Me.cboMascara.ListIndex = -1
                Me.txtTamano.Text = ""
                Me.txtTamano.Enabled = False
                Me.cboMascara.Enabled = False
                Me.cmdValores.Visible = False
                Me.chkValidarRepetidos.Value = 0
                Me.chkValidarRepetidos.Enabled = False
            Case "B"
                Me.cboMascara.ListIndex = -1
                Me.txtTamano.Text = ""
                Me.txtTamano.Enabled = False
                Me.cboMascara.Enabled = False
                Me.cmdValores.Visible = False
                Me.chkValidarRepetidos.Value = 0
                Me.chkValidarRepetidos.Enabled = False
        End Select

    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub chkInterfase_Click()
    On Error GoTo ErrManager
    
    If Me.chkInterfase.Value = 1 Then
        Me.txtPosicion.Enabled = True
        If Me.pnlInterfase.Enabled = True Then
            Me.txtPosicion.SetFocus
        End If
    Else
        Me.txtPosicion.Text = ""
        Me.txtPosicion.Enabled = False
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmBuscarCampos_Click()
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    If Me.cboNombreProducto.Text = "" Or Me.cboNombreProducto.ListIndex = -1 Then
        MsgBox "Debe seleccionar el producto a buscar.", vbInformation, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Set Me.proParametrosProducto = Nothing
    Set Me.proParametrosProducto = New colParametroProducto
    Set Me.proParametrosProducto.proConexion = Me.proConexion
    
    Me.proParametrosProducto.proProductNumber = Me.cboCodigoProducto.Text
    
    If Me.proParametrosProducto.metConsultarxProducto Then
        Call SubFPintarGrid
        Call cmdCancelar_Click
        Me.cmdNuevo.Enabled = True
        Me.grdCampos.Enabled = True
    Else
        MsgBox "Error al consultar los campos x producto.", vbCritical, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrManager
    
    'Llenar el combo de campos
    Call SubFLlenarComboCampos("C")
    
    Call grdCampos_Click

    Me.cmdGuardar.Enabled = False
    Me.cmdNuevo.Enabled = True
    Me.cmdCancelar.Enabled = False
    
    Me.pnlCampo.Enabled = False
    Me.pnlInterfase.Enabled = False
    
    Me.grdCampos.Enabled = True
    Me.cboCampos.Enabled = True
    Me.cmdValores.Visible = False
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdEliminar_Click()
    Dim varParametrosxProducto As claParametroProducto
    On Error GoTo ErrManager
    
        If MsgBox("Desea eliminar el campo [" & Me.cboCampos.Text & "]?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
        
        Set varParametrosxProducto = New claParametroProducto
        Set varParametrosxProducto.proConexion = Me.proConexion
        
        varParametrosxProducto.proProductNumber = Me.cboCodigoProducto.Text
        varParametrosxProducto.proCampo = Me.cboCampos.Text
        
        If varParametrosxProducto.MetEliminar Then
            Set Me.proParametrosProducto = Nothing
            Set Me.proParametrosProducto = New colParametroProducto
            Set Me.proParametrosProducto.proConexion = Me.proConexion
            Me.proParametrosProducto.proProductNumber = Me.cboCodigoProducto.Text
            If Me.proParametrosProducto.metConsultarxProducto Then
                Call SubFPintarGrid
                Me.pnlCampo.Enabled = False
                Me.pnlInterfase.Enabled = False
                Me.cmdGuardar.Enabled = False
                Me.cmdEliminar.Enabled = False
                Me.cmdNuevo.Enabled = True
                Me.grdCampos.Enabled = True
                Me.cmdCancelar.Enabled = False
                Me.cmdNuevo.SetFocus
            Else
                MsgBox "Error al consultar los parametros del producto elegido.", vbCritical, App.Title
                Exit Sub
            End If
        Else
            MsgBox "Error al eliminar el parámetro.", vbCritical, App.Title
            Exit Sub
        End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdGuardar_Click()
    Dim varParametrosxProducto As claParametroProducto
    On Error GoTo ErrManager
    
    If FunFValidarInformacion Then
    
        Set varParametrosxProducto = New claParametroProducto
        Set varParametrosxProducto.proConexion = Me.proConexion
        
        varParametrosxProducto.proProductNumber = Me.cboCodigoProducto.Text
        varParametrosxProducto.proCampo = Me.cboCampos.Text
        varParametrosxProducto.proEtiqueta = Trim(Me.txtEtiqueta.Text)
        varParametrosxProducto.proTipo = Me.cboCodigoTipo.Text
        varParametrosxProducto.proTamaño = Val(Me.txtTamano.Text)
        varParametrosxProducto.proIDInterfase = Me.chkInterfase.Value
        varParametrosxProducto.proPosicionInterfase = Val(Me.txtPosicion.Text)
        varParametrosxProducto.proMascara = Me.cboCodigoMascara.Text
        varParametrosxProducto.proObligatorioVenta = Me.chkObligatorioVenta.Value
        varParametrosxProducto.proObligatorioAtencion = Me.chkObligatorioAtencion.Value
        varParametrosxProducto.proObligatorioOT = Me.chkObligatorioOT.Value
        varParametrosxProducto.proValidarRepetidos = Val(Me.chkValidarRepetidos.Value)
        varParametrosxProducto.proCampoPadre = IIf(proCampoPadre = "0", varParametrosxProducto.proCampo, proCampoPadre)

        'Validar si es una insercion o una actualizacion
        If Me.cboCampos.Enabled = True Then 'Insercion
            If varParametrosxProducto.MetInsertar Then
                Set Me.proParametrosProducto = Nothing
                Set Me.proParametrosProducto = New colParametroProducto
                Set Me.proParametrosProducto.proConexion = Me.proConexion
                Me.proParametrosProducto.proProductNumber = Me.cboCodigoProducto.Text
                If Me.proParametrosProducto.metConsultarxProducto Then
                    Call SubFPintarGrid
                    Me.pnlCampo.Enabled = False
                    Me.pnlInterfase.Enabled = False
                    Me.cmdGuardar.Enabled = False
                    Me.cmdEliminar.Enabled = False
                    Me.cmdNuevo.Enabled = True
                    Me.grdCampos.Enabled = True
                    Me.cmdCancelar.Enabled = False
                    Me.cmdValores.Visible = False
                    Me.cmdNuevo.SetFocus
                    
                    MsgBox "El parámetro se insertó exitosamente.", vbInformation, App.Title
                    'Si el tipo del campo es lista se deben asignar los valores
                    
                    Exit Sub
                Else
                    MsgBox "Error al consultar los parametros del producto elegido.", vbCritical, App.Title
                    Exit Sub
                End If
            Else
                MsgBox "Error al insertar el parámetro para el producto.", vbCritical, App.Title
                Exit Sub
            End If
        Else    'Actualizacion
            If varParametrosxProducto.MetActualizar Then
                Set Me.proParametrosProducto = Nothing
                Set Me.proParametrosProducto = New colParametroProducto
                Set Me.proParametrosProducto.proConexion = Me.proConexion
                Me.proParametrosProducto.proProductNumber = Me.cboCodigoProducto.Text
                If Me.proParametrosProducto.metConsultarxProducto Then
                    Call SubFPintarGrid
                    Me.pnlCampo.Enabled = False
                    Me.pnlInterfase.Enabled = False
                    Me.cmdGuardar.Enabled = False
                    Me.cmdEliminar.Enabled = False
                    Me.cmdNuevo.Enabled = True
                    Me.grdCampos.Enabled = True
                    Me.cmdCancelar.Enabled = False
                    Me.cmdValores.Visible = False
                    Me.cmdNuevo.SetFocus
                    MsgBox "El parámetro se actualizó exitosamente.", vbInformation, App.Title
                    Unload Me
                    Exit Sub
                Else
                    MsgBox "Error al consultar los parámetros del producto elegido.", vbCritical, App.Title
                    Exit Sub
                End If
            Else
                MsgBox "Error al actualizar el parámetro para el producto.", vbCritical, App.Title
                Exit Sub
            End If
        End If
    End If

    Exit Sub
ErrManager:
    SubGMuestraError
    
End Sub

Private Sub cmdInsertar_Click()
    Set Me.proParametrosProducto = Nothing
    Set Me.proParametrosProducto = New colParametroProducto
    Set Me.proParametrosProducto.proConexion = Me.proConexion
    Me.proParametrosProducto.proProductNumber = Me.cboCodigoProducto.Text

End Sub

Private Sub cmdModificar_Click()
    On Error GoTo ErrManager
    
    Me.pnlCampo.Enabled = True
    Me.pnlInterfase.Enabled = True
    Me.cmdGuardar.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.cmdModificar.Enabled = False
    Me.cmdNuevo.Enabled = False
    Me.cmdEliminar.Enabled = False
    Me.cmdValores.Enabled = True
    
    Me.grdCampos.Enabled = False
    Me.cboCampos.Enabled = False
    Me.txtEtiqueta.SetFocus
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo ErrManager
    
    Me.cboCampos.ListIndex = -1
    Me.cboMascara.ListIndex = -1
    Me.cboTipo.ListIndex = -1
    Me.txtEtiqueta.Text = ""
    Me.txtTamano.Text = ""
    Me.chkObligatorioVenta.Value = 0
    Me.chkObligatorioAtencion.Value = 0
    Me.chkObligatorioOT.Value = 0
    Me.chkInterfase.Value = 0
    Me.chkValidarRepetidos.Value = 0
    Me.chkValidarRepetidos.Enabled = False
    Me.txtPosicion.Text = ""
    
    Me.pnlCampo.Enabled = True
    Me.pnlInterfase.Enabled = True
    Me.cmdGuardar.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.cmdNuevo.Enabled = False
    Me.cmdModificar.Enabled = False
    Me.cmdEliminar.Enabled = False
    Me.cboCampos.Enabled = True
    
    Me.grdCampos.Enabled = False
    
    'Llenar el combo de campos
    Call SubFLlenarComboCampos("N")
    
    'Me.cboCampos.SetFocus
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdValores_Click()
    On Error GoTo ErrManager
    
    'Si el tipo del campo es lista se deben asignar los valores
    Set frmAsignarValores.proConexion = Me.proConexion
    Set frmAsignarValores.proParametroProducto = Me.proParametrosProducto.Item(Me.grdCampos.Row)
    frmAsignarValores.Show vbModal

    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrManager
    
    Call SubFInicializarPantalla

    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarPantalla()
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    Me.cmdNuevo.Enabled = False
    Me.cmdGuardar.Enabled = False
    Me.cmdCancelar.Enabled = False
    Me.cmdModificar.Enabled = False
    Me.cmdEliminar.Enabled = False
    
    Me.pnlCampo.Enabled = False
    Me.pnlInterfase.Enabled = False
    
    'Consultar los productos activos y llenar los combos
    Call SubFLlenarComboProductos
    
    'Inicializar el grid de campos
    Call SubFInicializarGrid
        
    'Llenar el combo de tipos de campos
    Call SubFLlenarComboTipos
    
    'Llenar el combo de mascaras
    Call SubFLlenarComboMascaras
    
    Screen.MousePointer = 0
    
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboProductos()
    On Error GoTo ErrManager
    
    Set Me.proProducto = New colProductMaster
    
    Set Me.proProducto.proConexion = Me.proConexion
    
    If Me.proProducto.MetConsultar Then
        Call SubFPintarComboProductos
    Else
        MsgBox "Error al consultar los productos.", vbCritical, App.Title
        Exit Sub
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarComboProductos()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.cboCodigoProducto.Clear
    Me.cboNombreProducto.Clear
    
    For varContador = 1 To Me.proProducto.Count
        Me.cboNombreProducto.AddItem Me.proProducto.Item(varContador).proDescription
        Me.cboCodigoProducto.AddItem Me.proProducto.Item(varContador).proProductNumber
    Next varContador
    
    Me.cboCodigoProducto.ListIndex = -1
    Me.cboNombreProducto.ListIndex = -1
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGrid()
    On Error GoTo ErrManager
    
    With Me.grdCampos
        .Rows = 1
        .Cols = 12
        
        .Row = 0
        .Col = 0
        .ColWidth(0) = 1470
        .CellAlignment = 4
        .TextMatrix(0, 0) = "Product Number"
        
        .Col = 1
        .ColWidth(1) = 1005
        .CellAlignment = 4
        .TextMatrix(0, 1) = "Campo"
        
        .Col = 2
        .ColWidth(2) = 1005
        .CellAlignment = 4
        .TextMatrix(0, 2) = "Etiqueta"
        
        .Col = 3
        .ColWidth(3) = 1005
        .CellAlignment = 4
        .TextMatrix(0, 3) = "Tipo"
    
        .Col = 4
        .ColWidth(4) = 1005
        .CellAlignment = 4
        .TextMatrix(0, 4) = "Tamaño"
    
        .Col = 5
        .ColWidth(5) = 1575
        .CellAlignment = 4
        .TextMatrix(0, 5) = "Interfase Billing"
        
        .Col = 6
        .ColWidth(6) = 1590
        .CellAlignment = 4
        .TextMatrix(0, 6) = "Posición Interfase"
        
        .Col = 7
        .ColWidth(7) = 1005
        .CellAlignment = 4
        .TextMatrix(0, 7) = "Mascara"
        
        .Col = 8
        .ColWidth(8) = 1605
        .CellAlignment = 4
        .TextMatrix(0, 8) = "Obligatorio Venta"
        
        .Col = 9
        .ColWidth(9) = 1605
        .CellAlignment = 4
        .TextMatrix(0, 9) = "Obligatorio Atencion"
        
        .Col = 10
        .ColWidth(10) = 1605
        .CellAlignment = 4
        .TextMatrix(0, 10) = "Obligatorio OT"
        
        .Col = 11
        .ColWidth(11) = 1605
        .CellAlignment = 4
        .TextMatrix(0, 11) = "Validar Repetidos"
        .Row = 0
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Public Sub SubFPintarGrid()
    Dim varContador As Integer
    Dim varTipo As String
    Dim varInterfase  As String
    Dim varMascara As String
    Dim varObligatorioVenta As String
    Dim varObligatorioAtencion As String
    Dim varObligatorioOT As String
    Dim varValidarRepetidos As String
    On Error GoTo ErrManager
    
    Me.grdCampos.Rows = 1
    For varContador = 1 To Me.proParametrosProducto.Count
        Select Case Me.proParametrosProducto.Item(varContador).proTipo
            Case "T"
                varTipo = "Texto"
            Case "F"
                varTipo = "Fecha"
            Case "L"
                varTipo = "Lista"
            Case "B"
                varTipo = "Booleano"
        End Select
        
        If Me.proParametrosProducto.Item(varContador).proIDInterfase = True Then
            varInterfase = "Si"
        Else
            varInterfase = "No"
        End If
        
        Select Case Me.proParametrosProducto.Item(varContador).proMascara
            Case "N"
                varMascara = "Numérico"
            Case "A"
                varMascara = "AlfaNumérico"
            Case Else
                varMascara = ""
        End Select
        
        If Me.proParametrosProducto.Item(varContador).proObligatorioVenta = True Then
            varObligatorioVenta = "Si"
        Else
            varObligatorioVenta = "No"
        End If
        
        If Me.proParametrosProducto.Item(varContador).proObligatorioAtencion = True Then
            varObligatorioAtencion = "Si"
        Else
            varObligatorioAtencion = "No"
        End If
        
        If Me.proParametrosProducto.Item(varContador).proObligatorioOT = True Then
            varObligatorioOT = "Si"
        Else
            varObligatorioOT = "No"
        End If
        
        If Me.proParametrosProducto.Item(varContador).proValidarRepetidos = "True" Or Me.proParametrosProducto.Item(varContador).proValidarRepetidos = "1" Then
            varValidarRepetidos = "Si"
        Else
            varValidarRepetidos = "No"
        End If
        
        Me.grdCampos.AddItem Me.proParametrosProducto.Item(varContador).proProductNumber & vbTab & _
                             Me.proParametrosProducto.Item(varContador).proCampo & vbTab & _
                             Me.proParametrosProducto.Item(varContador).proEtiqueta & vbTab & _
                             varTipo & vbTab & _
                             Me.proParametrosProducto.Item(varContador).proTamaño & vbTab & _
                             varInterfase & vbTab & _
                             Me.proParametrosProducto.Item(varContador).proPosicionInterfase & vbTab & _
                             varMascara & vbTab & _
                             varObligatorioVenta & vbTab & _
                             varObligatorioAtencion & vbTab & _
                             varObligatorioOT & vbTab & _
                             varValidarRepetidos
                             
    Next varContador
    
    Me.grdCampos.Col = 0
    Me.grdCampos.Row = 0
    Me.cmdEliminar.Enabled = False
    Me.cmdModificar.Enabled = False
    
    Me.cboCampos.ListIndex = -1
    Me.cboMascara.ListIndex = -1
    Me.cboTipo.ListIndex = -1
    Me.txtEtiqueta.Text = ""
    Me.txtTamano.Text = ""
    Me.chkObligatorioVenta.Value = 0
    Me.chkObligatorioAtencion.Value = 0
    Me.chkObligatorioOT.Value = 0
    Me.chkInterfase.Value = 0
    Me.chkValidarRepetidos.Value = 0
    Me.txtPosicion.Text = ""
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Public Sub BuscarRegistroEnGrid()
    Dim i As Integer
    For i = 0 To grdCampos.Rows - 1
        If Trim(grdCampos.TextMatrix(i, 0)) = proProductNumber And Trim(grdCampos.TextMatrix(i, 1)) = proCampo Then
            grdCampos.Row = i
            Exit For
        End If
    Next
End Sub

Sub limpiarCamposPanel()
        Me.cboCampos.ListIndex = -1
        Me.cboMascara.ListIndex = -1
        Me.cboTipo.ListIndex = -1
        Me.txtEtiqueta.Text = ""
        Me.txtTamano.Text = ""
        Me.chkObligatorioVenta.Value = 0
        Me.chkObligatorioAtencion.Value = 0
        Me.chkObligatorioOT.Value = 0
        Me.chkInterfase.Value = 0
        Me.chkValidarRepetidos = 0
        Me.txtPosicion.Text = ""
End Sub
Public Sub grdCampos_Click()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    If Me.grdCampos.Row = 0 Then
        Call limpiarCamposPanel
    Else
        Me.cmdEliminar.Enabled = True
        Me.cmdModificar.Enabled = True
        Me.cmdNuevo.Enabled = True
        If Me.proParametrosProducto.Item(Me.grdCampos.Row).proTipo = "L" Then
            Me.cmdValores.Visible = True
            Me.cmdValores.Enabled = False
        Else
            Me.cmdValores.Visible = False
        End If
        
        Call SubFLlenarComboCampos("C")
        'Asignar valores a los objetos en la parte inferior
        For varContador = 0 To Me.cboCampos.ListCount - 1
            Me.cboCampos.ListIndex = varContador
            If Me.cboCampos.Text = Me.proParametrosProducto.Item(Me.grdCampos.Row).proCampo Then
                Exit For
            End If
        Next varContador
        
        Me.txtEtiqueta.Text = Me.proParametrosProducto.Item(Me.grdCampos.Row).proEtiqueta
        
        For varContador = 0 To Me.cboCodigoTipo.ListCount - 1
            Me.cboTipo.ListIndex = varContador
            If Me.cboCodigoTipo.Text = Me.proParametrosProducto.Item(Me.grdCampos.Row).proTipo Then
                Exit For
            End If
        Next varContador
        
        If Trim(Me.proParametrosProducto.Item(Me.grdCampos.Row).proMascara) = "" Then
            Me.cboCodigoMascara.ListIndex = -1
        Else
            For varContador = 0 To Me.cboCodigoMascara.ListCount - 1
                Me.cboMascara.ListIndex = varContador
                If Me.cboCodigoMascara.Text = Me.proParametrosProducto.Item(Me.grdCampos.Row).proMascara Then
                    Exit For
                End If
            Next varContador
        End If
        Me.txtTamano.Text = Me.proParametrosProducto.Item(Me.grdCampos.Row).proTamaño
        
        If Me.proParametrosProducto.Item(Me.grdCampos.Row).proObligatorioVenta = True Then
            Me.chkObligatorioVenta.Value = 1
        Else
            Me.chkObligatorioVenta.Value = 0
        End If
        
        If Me.proParametrosProducto.Item(Me.grdCampos.Row).proObligatorioAtencion = True Then
            Me.chkObligatorioAtencion.Value = 1
        Else
            Me.chkObligatorioAtencion.Value = 0
        End If
        
        If Me.proParametrosProducto.Item(Me.grdCampos.Row).proObligatorioOT = True Then
            Me.chkObligatorioOT.Value = 1
        Else
            Me.chkObligatorioOT.Value = 0
        End If
        
        If Me.proParametrosProducto.Item(Me.grdCampos.Row).proIDInterfase = True Then
            Me.chkInterfase.Value = 1
        Else
            Me.chkInterfase.Value = 0
        End If
        
        If Me.proParametrosProducto.Item(Me.grdCampos.Row).proValidarRepetidos = "True" Or Trim(Me.proParametrosProducto.Item(Me.grdCampos.Row).proValidarRepetidos) = "1" Then
            Me.chkValidarRepetidos.Value = 1
        Else
            Me.chkValidarRepetidos.Value = 0
        End If
        
        If Me.proParametrosProducto.Item(Me.grdCampos.Row).proTipo = "L" Then
            Me.chkValidarRepetidos.Enabled = True
        Else
            Me.chkValidarRepetidos.Enabled = False
        End If
        Me.txtPosicion.Text = Me.proParametrosProducto.Item(Me.grdCampos.Row).proPosicionInterfase
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Public Sub SubFLlenarComboCampos(parOrigen As String)
    Dim varContador As Integer
    Dim varContadorAux As Integer
    Dim varEncontro As Boolean
    On Error GoTo ErrManager
    
    Me.cboCampos.Clear
    If parOrigen = "N" Then
        For varContador = 1 To 40
        
            varEncontro = False
            For varContadorAux = 1 To Me.proParametrosProducto.Count
                If Trim(Me.proParametrosProducto.Item(varContadorAux).proCampo) = ("vchUser" & CStr(varContador)) Then
                    varEncontro = True
                End If
            Next varContadorAux
            
            If Not varEncontro Then
                Me.cboCampos.AddItem "vchUser" & CStr(varContador)
            End If
        Next varContador
    Else
        For varContador = 1 To 40
            Me.cboCampos.AddItem "vchUser" & CStr(varContador)
        Next varContador
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboTipos()
    On Error GoTo ErrManager
           
        Me.cboTipo.Clear
        Me.cboCodigoTipo.Clear
        
        Me.cboTipo.AddItem "Texto"
        Me.cboCodigoTipo.AddItem "T"
       
        Me.cboTipo.AddItem "Fecha"
        Me.cboCodigoTipo.AddItem "F"
        
        Me.cboTipo.AddItem "Lista"
        Me.cboCodigoTipo.AddItem "L"
        
        Me.cboTipo.AddItem "Booleano"
        Me.cboCodigoTipo.AddItem "B"
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboMascaras()
    On Error GoTo ErrManager
    
    Me.cboMascara.Clear
    Me.cboCodigoMascara.Clear
    
    Me.cboMascara.AddItem "Numérico"
    Me.cboCodigoMascara.AddItem "N"
    
    Me.cboMascara.AddItem "AlfaNumérico"
    Me.cboCodigoMascara.AddItem "A"
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtEtiqueta_GotFocus()
    On Error GoTo ErrManager
    
    Me.txtEtiqueta.SelStart = 0
    Me.txtEtiqueta.SelLength = Len(Me.txtEtiqueta.Text)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtEtiqueta_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
    KeyAscii = FunGLeeAlfaNumerico(KeyAscii, 1)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtPosicion_GotFocus()
    On Error GoTo ErrManager
    
    Me.txtPosicion.SelStart = 0
    Me.txtPosicion.SelLength = Len(Me.txtPosicion.Text)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtPosicion_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtTamano_GotFocus()
    On Error GoTo ErrManager
    
    Me.txtTamano.SelStart = 0
    Me.txtTamano.SelLength = Len(Me.txtTamano.Text)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtTamano_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Function FunFValidarInformacion() As Boolean
    Dim varContador As Integer
    Dim varExiste As Boolean
    On Error GoTo ErrManager
    
    If Me.cboCampos.ListIndex = -1 Or Trim(Me.cboCampos.Text) = "" Then
        MsgBox "Debe seleccionar el campo que desea asignar.", vbInformation, App.Title
        FunFValidarInformacion = False
        Exit Function
    End If
    
    If Trim(Me.txtEtiqueta.Text) = "" Then
        MsgBox "Debe digitar la etiqueta que se le asignará al campo.", vbInformation, App.Title
        FunFValidarInformacion = False
        Exit Function
    End If
    
    If Me.cboTipo.ListIndex = -1 Or Trim(Me.cboTipo.Text) = "" Then
        MsgBox "Debe seleccionar el tipo de campo.", vbInformation, App.Title
        FunFValidarInformacion = False
        Exit Function
    End If
    
    If Me.cboCodigoTipo.Text = "T" Then
        If Me.cboMascara.ListIndex = -1 Or Trim(Me.cboMascara.Text) = "" Then
            MsgBox "Debe seleccionar la Mascara para los campos tipo texto.", vbInformation, App.Title
            FunFValidarInformacion = False
            Exit Function
        End If
        Set proclaParametrosProducto = New claParametroProducto
    If Me.cboCodigoTipo.Text <> "L" And Not Me.cboCampos.Enabled Then
        If Me.proParametrosProducto.Item(Me.grdCampos.Row).proTipo = "L" Then
            proclaParametrosProducto.proCampo = proParametrosProducto.Item(Me.grdCampos.Row).proCampo
            proclaParametrosProducto.proCampoPadre = proParametrosProducto.Item(Me.grdCampos.Row).proCampoPadre
            proclaParametrosProducto.proProductNumber = proParametrosProducto.Item(Me.grdCampos.Row).proProductNumber
            Set proclaParametrosProducto.proConexion = Me.proConexion
            If proclaParametrosProducto.MetTieneHijos Then
                MsgBox "Este parámetro tiene hijos. debe eliminarlos antes de cambiar el tipo", vbInformation, App.Title
                FunFValidarInformacion = False
                Exit Function
            End If
        End If
        If Me.cboMascara.ListIndex = -1 Or Trim(Me.cboMascara.Text) = "" Then
            MsgBox "Debe seleccionar la Mascara para los campos tipo texto.", vbInformation, App.Title
            FunFValidarInformacion = False
            Exit Function
        End If
    End If
        
        If Trim(Me.txtTamano.Text) = "" Then
            MsgBox "Debe selecionar el tamaño para los campos tipo texto.", vbInformation, App.Title
            FunFValidarInformacion = False
            Exit Function
        End If
    End If
    
    If Me.chkInterfase.Value = 1 Then
        If Trim(Me.txtPosicion.Text) = "" Then
            MsgBox "Debe digitar la posición del campo en la interfase.", vbInformation, App.Title
            FunFValidarInformacion = False
            Exit Function
        End If
        
        'Recorrer la coleccion de campos por producto para verificar que no existan
        'dos campos con el mismo orden en la interfase}
        If Me.cboCampos.Enabled = True Then
            varExiste = False
            For varContador = 1 To Me.proParametrosProducto.Count
                If Val(Me.txtPosicion.Text) = Val(Me.proParametrosProducto.Item(varContador).proPosicionInterfase) Then
                    varExiste = True
                    Exit For
                End If
            Next varContador
            
            If varExiste Then
                MsgBox "Ya existe un campo el cual será enviado a la interfase en la posición [" & Me.txtPosicion.Text & "]", vbInformation, App.Title
                FunFValidarInformacion = False
                Exit Function
            End If
        Else
            If Me.proParametrosProducto.Item(Me.grdCampos.Row).proPosicionInterfase <> Me.txtPosicion.Text Then
                For varContador = 1 To Me.proParametrosProducto.Count
                    If Val(Me.txtPosicion.Text) = Val(Me.proParametrosProducto.Item(varContador).proPosicionInterfase) And varContador <> Me.grdCampos.Row Then
                        varExiste = True
                        Exit For
                    End If
                Next varContador
                
                If varExiste Then
                    MsgBox "Ya existe un campo el cual será enviado a la interfase en la posición [" & Me.txtPosicion.Text & "]", vbInformation, App.Title
                    Me.txtPosicion.Text = Me.proParametrosProducto.Item(Me.grdCampos.Row).proPosicionInterfase
                    FunFValidarInformacion = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    FunFValidarInformacion = True
    Exit Function
ErrManager:
    SubGMuestraError
    
End Function
