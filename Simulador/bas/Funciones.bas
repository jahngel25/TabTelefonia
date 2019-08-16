Attribute VB_Name = "Module1"
Option Explicit

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
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 44 Then
       FunGLeeDecimales = 0
    Else
       Exit Function
    End If
    
    'Caracteres de excepción ( no hay que validar )
    If KeyAscii = 8 Then
        FunGLeeDecimales = 8
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
        
        'Construye la cadena con el valor si fuera incluido
        varTexto = Left(parTexto, parTexto.SelStart)  'Izquierda
        varTexto = varTexto & Chr(KeyAscii) 'Nuevo caracter
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


