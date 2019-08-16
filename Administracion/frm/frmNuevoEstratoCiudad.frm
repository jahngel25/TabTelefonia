VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmNuevoEstratoCiudad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Estratos Por Ciudad"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4065
   Icon            =   "frmNuevoEstratoCiudad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   2212
      TabIndex        =   5
      ToolTipText     =   "Cancelar los cambios"
      Top             =   1425
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   390
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Guardar la configuración"
      Top             =   1440
      Width           =   1215
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   337
      TabIndex        =   3
      Top             =   900
      Width           =   690
      Caption         =   "Estrato"
      Size            =   "1217;423"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   240
      Left            =   337
      TabIndex        =   2
      Top             =   375
      Width           =   690
      Caption         =   "Ciudad"
      Size            =   "1217;423"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtEstrato 
      Height          =   390
      Left            =   1087
      TabIndex        =   1
      ToolTipText     =   "Nombre del estrato"
      Top             =   825
      Width           =   2640
      VariousPropertyBits=   746604571
      MaxLength       =   50
      Size            =   "4657;688"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbCiudad 
      Height          =   390
      Left            =   1087
      TabIndex        =   0
      ToolTipText     =   "Ciudad del estrato"
      Top             =   300
      Width           =   2640
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4657;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbiTelefoniaCiudadId 
      Height          =   240
      Left            =   1425
      TabIndex        =   6
      Top             =   900
      Visible         =   0   'False
      Width           =   1815
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3201;423"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmNuevoEstratoCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proEstratoCiudad As claEstratoCiudad
Public proEstratoCiudadCol As colEstratoCiudad
Public proConexion As ADODB.Connection

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Function Trim2(varcadena As String) As String
    Trim2 = ""
    Dim varcontador As Integer
    For varcontador = 1 To Len(varcadena)
        If Mid(varcadena, varcontador, 1) <> " " Then
            Trim2 = Trim2 & Mid(varcadena, varcontador, 1)
        End If
    Next
End Function
Private Sub cmdGuardar_Click()
    Dim varcontador As Boolean
    Dim i As Integer
    Set proEstratoCiudad = New claEstratoCiudad
    Set proEstratoCiudadCol = New colEstratoCiudad
    Set proEstratoCiudadCol.proConexion = Me.proConexion
    cmbiTelefoniaCiudadId.ListIndex = cmbCiudad.ListIndex
    
    varcontador = proEstratoCiudadCol.FunGConsulta(0, 1)
    varcontador = proEstratoCiudadCol.FunGConsulta(0, 0)
    txtEstrato.Text = Trim(txtEstrato)
    If cmbCiudad.ListIndex = 0 Then
        MsgBox "Seleccione una ciudad", vbInformation, "Seleccione una ciudad"
    ElseIf txtEstrato.Text = "" Then
        MsgBox "Ingrese el estrato", vbInformation, "Ingrese el estrato"
    Else
        For i = 1 To proEstratoCiudadCol.Count
            If proEstratoCiudadCol.Item(i).proCiudadId = cmbiTelefoniaCiudadId.Value _
                And UCase(Trim2(proEstratoCiudadCol.Item(i).proNombreEstrato)) = UCase(Trim2(txtEstrato.Text)) Then
                MsgBox "Ya existe el estrato " & txtEstrato.Text & " para la ciudad de " & cmbCiudad & ", verifique su estado", vbInformation
                txtEstrato.Text = ""
                Exit Sub
            End If
        Next i
        proEstratoCiudad.proCiudadId = cmbiTelefoniaCiudadId.Value
        proEstratoCiudad.proNombreEstrato = txtEstrato.Text
        Set proEstratoCiudad.proConexion = Me.proConexion
        varcontador = proEstratoCiudad.FunGInsertar
        MsgBox "Estrato creado con éxito", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim objNewMember As colCiudadOnyx
    Set objNewMember = New colCiudadOnyx
    Set objNewMember.proConexion = Me.proConexion
    objNewMember.FunGConsulta
    Call FunGLlenarCombosCiudad(cmbiTelefoniaCiudadId, cmbCiudad, objNewMember, "Seleccione una ciudad")
    Set proEstratoCiudad = New claEstratoCiudad
End Sub

Private Sub txtEstrato_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 45 Or KeyAscii = 32 Then
'        If KeyAscii = 32 Then 'Espacio a evaluar
'            If Right(txtEstrato.Text, 1) = " " Then Exit Sub
'        End If
'    End If

End Sub
