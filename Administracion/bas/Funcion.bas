Attribute VB_Name = "Function"
Sub SubGMuestraError()
'***********************************************************
'   OBJETIVO:  Manejo de Errores Centralizado, captura el
'              error y lo despliega
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 21/12/2000
'***********************************************************
    MsgBox "[" & Trim(Str(Err.Number)) & "] - " & _
        Err.Description, vbOKOnly + vbInformation, App.Title
End Sub

Function FunGLeeDecimales(KeyAscii As Integer, parTexto As TextBox) As Integer
'************************************************************************
'*  OBJETIVOS :  Lee la cantidad indicada de enteros y decimales
'************************************************************************
'*  PARAMETROS:
'*      KeyAscii                Ascii de la tecla oprimida
'*      ParTexto                Texto a validar
'*
'*  RESULTADOS:
'*      #                           Tecla Válida
'*      0                           Tecla Inválida
'*************************************************************************
'*  SONDA de Colombia
'*  Autor: Raúl Cruz A.
'*  Fecha: 06 / 02 / 2001
'***********************************************************************
Dim varEnteros As Integer
Dim varDecimales As Integer
Dim varEnterosActuales As Integer
Dim varDecimalesActuales As Integer
Dim varValorMaximo As Double
Dim varPosicionPunto As Integer
Dim varValor As Double
Dim varTexto As String
Dim varTextoSeleccion As String
On Error GoTo ErrorManager



    'valida que los caracteres sean numéricos
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Then
       FunGLeeDecimales = 0
    Else
       Exit Function
    End If
    
   
    parTexto = Trim(parTexto)
    varTextoSeleccion = Left(parTexto, parTexto.SelStart) + Right(parTexto, Len(parTexto) - (parTexto.SelStart + parTexto.SelLength))
    
    If Len(Trim(parTexto.Tag)) > 0 Then
        'Toma el número de enteros y Decimales
        varPosicionPunto = InStr(parTexto.Tag, ".")
        If varPosicionPunto = 0 Then
            varPosicionPunto = InStr(parTexto.Tag, ",")
        End If
        If varPosicionPunto Then
            varDecimales = Len(Trim(parTexto.Tag)) - varPosicionPunto
            varEnteros = varPosicionPunto - 1
        Else
            varEnteros = Len(Trim(parTexto.Tag))
        End If
        
        'Caso en que el ascii es el del punto o la coma
        If Chr(KeyAscii) = "." Or Chr(KeyAscii) = "," Then
                'Averigua si debe recibir coma o punto
                If varDecimales = 0 Then Exit Function
                
                'Averigua si ya existia una coma o punto
                If InStr(varTextoSeleccion, ".") > 0 Or InStr(varTextoSeleccion, ",") Then Exit Function
                
                If Len(Trim(varTextoSeleccion)) = 0 Then
                        parTexto = "0"
                        parTexto.SelStart = 1
                End If
        End If
        
        'Busca el valor máximo a leer
        varValorMaximo = Val(Left(parTexto.Tag, varEnteros))
        If varDecimales Then
            varValorMaximo = varValorMaximo + Val("0." & Right(parTexto.Tag, varDecimales))
        End If
        
        ''Caracteres de excepción ( no hay que validar )
        'If KeyAscii = 8 Then
        '    FunGLeeDecimales = 8
        '    'Falta verificar si la eliminación no es del punto, djando el número decimal
        '    'sobrepasar el límite estipulado
        '    'Construye la cadena con el valor si fuera incluido
        '    varTexto = Left(parTexto, parTexto.SelStart)  'Izquierda
        '    varTexto = varTexto & Right(parTexto, Len(Trim(parTexto)) - (parTexto.SelStart + parTexto.SelLength))  'Derecha
        '    Exit Function
        'End If
        
        
        'Construye la cadena con el valor si fuera incluido
        varTexto = Left(parTexto, parTexto.SelStart)  'Izquierda
        If KeyAscii <> 8 Then
                varTexto = varTexto & Chr(KeyAscii)                 'Nuevo caracter
        Else
                If parTexto.SelLength = 0 Then
                        varTexto = Left(varTexto, Len(varTexto) - 1)
                End If
        End If
        varTexto = varTexto & Right(parTexto, Len(Trim(parTexto)) - (parTexto.SelStart + parTexto.SelLength))  'Derecha
        
        'Busca el valor de la cadena incluyendo el caracter
        varPosicionPunto = InStr(varTexto, ".")
        If varPosicionPunto = 0 Then
            varPosicionPunto = InStr(varTexto, ",")
        End If
        If varPosicionPunto Then
            varDecimalesActuales = Len(Trim(varTexto)) - varPosicionPunto
            varEnterosActuales = varPosicionPunto - 1
        Else
            varEnterosActuales = Len(Trim(varTexto))
        End If
        
        'No puede tener más decimales de los estipulados, ni enteros de los estipulados
        If varDecimalesActuales > varDecimales Then Exit Function
        If varEnterosActuales > varEnteros Then Exit Function
        
        'Construye el valor
        varValor = Val(Left(varTexto, varEnterosActuales))
        If varDecimalesActuales Then
            varValor = varValor + Val("0." & Right(varTexto, varDecimalesActuales))
        End If
        
        'Si el valor actual + el nuevo caracter supera al máximo,
        'Retorna 0
        If varValor > varValorMaximo Then Exit Function
    End If
    
    'Retorna el valor del KeyAscii
    FunGLeeDecimales = KeyAscii
    Exit Function

ErrorManager:
    SubGMuestraError
End Function

Function FunGChequeo(ByVal parNumero As String) As Integer
'************************************************************************
'*  OBJETIVOS :  Calcular el digito de chequeo
'************************************************************************
'*  PARAMETROS:
'*      ParNumero           Numero base del cálculo
'*  RESULTADOS:
'*      #                           Digito de Chequeo
'*************************************************************************
'*  ADVANCED SOLUTIONS * COLOMBIA -  SONDA S.A.
'*  Autor: Raúl Cruz A.
'*  Fecha: 06 / 01 /1998
'************************************************************************
Dim VarDigitos As String
Dim VarTotal As Integer
Dim varCuenta As Integer
Dim VarResiduo As Integer
Dim varPosicion As Integer

        On Error GoTo ErrorManager
                
        parNumero = Format(parNumero, "000000000000000")
        VarDigitos = "716759534743413729231917130703"
        varPosicion = 1
        For varCuenta = 1 To 15
                VarTotal = VarTotal + (Val(Mid(parNumero, varCuenta, 1)) * Val(Mid(VarDigitos, varPosicion, 2)))
                varPosicion = varPosicion + 2
        Next varCuenta
        VarResiduo = VarTotal Mod 11
        If VarResiduo = 0 Or VarResiduo = 1 Then
                FunGChequeo = VarResiduo
        Else
                FunGChequeo = 11 - VarResiduo
        End If
        
        Exit Function

ErrorManager:
        MsgBox Err.Description, vbInformation
        Screen.MousePointer = vbDefault
End Function

Sub SubGTamObjeto(ParObjeto As Object, ParW As Integer, ParH As Integer, Optional ParNumDiv As Variant)
'******************************************************************************************
'*  OBJETIVOS :  Llevar el objeto al tamaño deseado de manera
'*                        gradual
'******************************************************************************************
'*  PARAMETROS:
'*      ParObjeto           Control a cambiar
'*      ParW                 Ancho del Control
'*      ParH                  Alto del Control
'*      ParNumDiv         Factor de Aceleración
'******************************************************************************************
'*  ADVANCED SOLUTIONS * COLOMBIA -  SONDA S.A.
'*  Autor: Raúl Cruz A.
'*  Fecha: 30/12/1997
'*******************************************************************************************
Dim VarCuentaX As Integer
Dim VarCuentaY As Integer
Dim VarPasoX As Integer
Dim VarPasoY As Integer
Dim varCuenta As Integer
Dim VarDiv As Integer

On Error GoTo ErrorManager

        ' VarPaso contiene el número de pixeles necesarios para que el crecimiento se realize en 10 ó 15
        ' pasos
        VarDiv = 15
        
        If IsMissing(ParNumDiv) = False Then VarDiv = CInt(ParNumDiv)
        VarPasoX = Abs((ParW - ParObjeto.Width) / VarDiv)
        VarPasoY = Abs((ParH - ParObjeto.Height) / VarDiv)
        If ParW < ParObjeto.Width Then VarPasoX = -VarPasoX
        If ParH < ParObjeto.Height Then VarPasoY = -VarPasoY
        For varCuenta = 1 To VarDiv
                    ParObjeto.Width = ParObjeto.Width + VarPasoX
                    ParObjeto.Height = ParObjeto.Height + VarPasoY
                    ParObjeto.Refresh
        Next varCuenta
        ParObjeto.Width = ParW
        ParObjeto.Height = ParH
        Exit Sub
        
ErrorManager:
        MsgBox Err.Description, vbInformation
        Screen.MousePointer = vbDefault
End Sub

Function FunGNumValor(ByVal ParAscii As Integer, parTexto As TextBox, Optional ParValor) As Integer
'**********************************************************************************************************
'*  OBJETIVOS :  Función de Validación de teclas, que consiste en validar la tecla, siempre
'*                        y cuando el valor que produce el texto sea menor o igual al Parámetro
'**********************************************************************************************************
'*  PARAMETROS:
'*      ParAscii            Tecla oprimida
'*      ParTexto            Objeto a evaluar
'*      ParValor            Valor máximo, si no se indica se opta por tomar el valor del tag del texto
'*
'*  RESULTADOS:
'*      0                           Fecha Inválida
'*      1                           Fecha Válida, no superior al día actual
'**********************************************************************************************************
'*  SONDA de Colombia
'*  Autor: Raúl Cruz A.
'*  Fecha: 28 / 01 /1999
'***********************************************************************************************************
Dim varValorMaximo As Double
Dim VarValor1 As String
Dim VarValor2 As String
On Error GoTo ErrorManager

            If ParAscii = 8 Then FunGNumValor = ParAscii
            If IsNumeric(Chr(ParAscii)) = False Then Exit Function
            If IsMissing(ParValor) = True Then ' No viene tomar el valor máximo del Tag
                    If Val(parTexto.Tag) = 0 Then Exit Function ' Se desea que el valor no tenga restricción
                    varValorMaximo = Val(parTexto.Tag)
            Else ' Tomar el Valor Máximo del Parámetro
                    varValorMaximo = Val(ParValor)
            End If
            If parTexto.SelStart Then
                    If parTexto.SelStart = Len(parTexto.Text) Then
                           VarValor1 = parTexto.Text
                           VarValor2 = ""
                    Else
                            VarValor1 = Left(parTexto.Text, Len(parTexto.Text) - parTexto.SelStart)
                            VarValor2 = Right(parTexto.Text, Len(parTexto.Text) - (parTexto.SelStart + parTexto.SelLength))
                    End If
            Else
                    VarValor1 = ""
                    VarValor2 = Right(parTexto.Text, Len(parTexto.Text) - (parTexto.SelStart + parTexto.SelLength))
            End If
            If CDbl(varValorMaximo) < CDbl(Trim(Format(Trim(VarValor1), "General Number")) _
                    & Chr(ParAscii) & Trim(Format(Trim(VarValor2), "General Number"))) Then
                    FunGNumValor = 0
            Else
                    FunGNumValor = ParAscii
            End If
            Exit Function
            
ErrorManager:
        MsgBox Err.Description, vbInformation
        Screen.MousePointer = vbDefault
End Function

Function FunGFechaDMA(ParFecha As String) As String
'***********************************************************
'   OBJETIVO:  Toma una fecha en formato MM/DD/AAAA y la
'              convierte a DD/MM/AAAA
'************************************************************
'   PARAMETROS:  ParFecha       Fecvha en formato MM/DD/AAAA
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 28/12/2000
'***********************************************************
On Error GoTo ErrorManager

        If Trim(ParFecha) = "" Then Exit Function

        FunGFechaDMA = Mid(ParFecha, 4, 2) & "/" & _
                       Left(ParFecha, 2) & "/" & _
                       Right(ParFecha, 4)
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function
Function FunGFechaMDA(ParFecha As String) As String
'***********************************************************
'   OBJETIVO:  Toma una fecha en formato DD/MM/AAAA y la convierte a
'              MM/DD/AAAA
'************************************************************
'   PARAMETROS:  ParFecha       Fecvha en formato DD/MM/AAAA
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 28/12/2000
'***********************************************************
On Error GoTo ErrorManager

        If Trim(ParFecha) = "" Then Exit Function

        FunGFechaMDA = Mid(ParFecha, 4, 2) & "/" & _
                       Left(ParFecha, 2) & "/" & _
                       Right(ParFecha, 4)
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function
Function FunGLeeNumerico(ParAscii As Integer) As Integer
'***********************************************************
'   OBJETIVO:  Manejo de Errores Centralizado, captura el
'              error y lo despliega
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 28/12/2000
'***********************************************************
On Error GoTo ErrorManager

    If (ParAscii >= 48 And ParAscii <= 57) Or ParAscii = 8 Then
        FunGLeeNumerico = ParAscii
    End If
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Function FunGLeeDireccion(ParAscii As Integer, parControl As TextBox) As Integer
'***********************************************************
'   OBJETIVO:  Leer Caractéres Mayúsculas, Minúsculas, Guión,
'              No se permiten dobles, espacios
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 28/12/2000
'***********************************************************
Dim varCaracter As String
On Error GoTo ErrorManager

    If (ParAscii >= 48 And ParAscii <= 57) Or _
        (ParAscii >= 65 And ParAscii <= 90) Or _
         (ParAscii >= 97 And ParAscii <= 122) Or _
                ParAscii = 8 Or ParAscii = 45 Or ParAscii = 32 Then
        If ParAscii = 32 Then 'Espacio a evaluar
            If Right(parControl.Text, 1) = " " Then Exit Function
        End If
        FunGLeeDireccion = ParAscii
    End If
    
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Function FunGLeeAlfaNumerico(ParAscii As Integer, Optional parCaseSensitive As Variant) As Integer
'***********************************************************
'   OBJETIVO:  Manejo de Errores Centralizado, captura el
'              error y lo despliega
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 28/12/2000
'***********************************************************
Dim varCaracter As String
On Error GoTo ErrorManager

    'Convierte a Mayúsculas el caracter
    varCaracter = Chr(ParAscii)
    If IsMissing(parCaseSensitive) = False Then
        If parCaseSensitive = False Then
            varCaracter = UCase(varCaracter)
        End If
    Else
        varCaracter = UCase(varCaracter)
    End If
    ParAscii = Asc(varCaracter)
    If IsMissing(parCaseSensitive) = False Then
            If parCaseSensitive Then
                If (ParAscii >= 97 And ParAscii <= 122) Or (ParAscii >= 49 And ParAscii <= 57) Or (ParAscii >= 65 And ParAscii <= 90) Or _
                            ParAscii = 8 Or ParAscii = 32 Then
                    FunGLeeAlfaNumerico = ParAscii
                End If
            Else
                If (ParAscii >= 49 And ParAscii <= 57) Or (ParAscii >= 65 And ParAscii <= 90) Or _
                            ParAscii = 8 Or ParAscii = 32 Then
                    FunGLeeAlfaNumerico = ParAscii
                End If
            End If
    Else
        If (ParAscii >= 49 And ParAscii <= 57) Or (ParAscii >= 65 And ParAscii <= 90) Or _
                            ParAscii = 8 Or ParAscii = 32 Then
                    FunGLeeAlfaNumerico = ParAscii
                End If
    End If
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Function FunGVerificaDatos(ParamArray parControles() As Variant) As Boolean
'***********************************************************
'   OBJETIVO:  Encuentra si falta algún dato en los controles,
'              en caso contrario muestra un mensaje y retorna el
'              foco al control
'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 29/12/2000
'***********************************************************
Dim varContador As Integer
Dim varEncontro As Integer

On Error GoTo ErrorManager

    varContador = 0
    varEncontro = False
    While varContador <= UBound(parControles) And varEncontro = False
        If TypeOf parControles(varContador) Is TextBox Then
            If Len(Trim(parControles(varContador))) = 0 Then varEncontro = True
        End If
        If TypeOf parControles(varContador) Is ComboBox Then
            If parControles(varContador).ListIndex = -1 Then varEncontro = True
        End If
        If TypeOf parControles(varContador) Is DTPicker Then
            If IsNull(parControles(varContador).Value) = True Then varEncontro = True
        End If
        
        If varEncontro = False Then
            varContador = varContador + 1
        Else
            MsgBox parControles(varContador).Tag, vbOKOnly + vbInformation, App.Title
            parControles(varContador).SetFocus
            Exit Function
        End If
    Wend
    
    FunGVerificaDatos = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

