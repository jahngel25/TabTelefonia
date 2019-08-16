Attribute VB_Name = "Generales"
Option Explicit
Public proConexion As ADODB.Connection

Function FunGLlenarCombosCiudad(ByRef cmbConsultariTelefoniaCiudadId As Object, ByRef cmbConsultarCiudad As Object, objNewMember As colCiudadOnyx, parElemento As String, Optional ParValor As Long = 0)
    Dim i As Integer
    cmbConsultarCiudad.Clear
    cmbConsultariTelefoniaCiudadId.Clear
    cmbConsultarCiudad.AddItem parElemento, 0
    cmbConsultariTelefoniaCiudadId.AddItem "0", 0
    For i = 1 To objNewMember.Count
        cmbConsultarCiudad.AddItem objNewMember.Item(i).proNombre
        cmbConsultariTelefoniaCiudadId.AddItem objNewMember.Item(i).proCiudadId
        If ParValor <> 0 Then
          If ParValor = objNewMember.Item(i).proCiudadId Then cmbConsultariTelefoniaCiudadId.ListIndex = i
        End If
    Next
    If ParValor = 0 Then
      cmbConsultarCiudad.ListIndex = 0
      cmbConsultariTelefoniaCiudadId.ListIndex = 0
    Else
        cmbConsultarCiudad.ListIndex = cmbConsultariTelefoniaCiudadId.ListIndex
    End If
End Function

