' Test simple para verificar funciones de CONDOR
Option Explicit

' Test MapPropKey
Function TestMapPropKey()
    Dim result
    result = "ancho -> " & MapPropKey("ancho") & vbCrLf
    result = result & "width -> " & MapPropKey("width") & vbCrLf
    result = result & "altura -> " & MapPropKey("altura") & vbCrLf
    TestMapPropKey = result
End Function

' Test NormalizeEnumToken
Function TestNormalizeEnumToken()
    Dim result
    result = "verdadero -> " & NormalizeEnumToken("verdadero", "Boolean") & vbCrLf
    result = result & "true -> " & NormalizeEnumToken("true", "Boolean") & vbCrLf
    result = result & "falso -> " & NormalizeEnumToken("falso", "Boolean") & vbCrLf
    TestNormalizeEnumToken = result
End Function

' Test ConvertColorToLong
Function TestConvertColorToLong()
    Dim result
    result = "#FF0000 -> " & ConvertColorToLong("#FF0000") & vbCrLf
    result = result & "#00FF00 -> " & ConvertColorToLong("#00FF00") & vbCrLf
    result = result & "#0000FF -> " & ConvertColorToLong("#0000FF") & vbCrLf
    TestConvertColorToLong = result
End Function

' Test MapEventToProperty
Function TestMapEventToProperty()
    Dim result
    result = "clic -> " & MapEventToProperty("clic") & vbCrLf
    result = result & "click -> " & MapEventToProperty("click") & vbCrLf
    result = result & "cargar -> " & MapEventToProperty("cargar") & vbCrLf
    result = result & "load -> " & MapEventToProperty("load") & vbCrLf
    TestMapEventToProperty = result
End Function

' Función principal
Sub Main()
    WScript.Echo "=== PRUEBAS DE FUNCIONES CONDOR ==="
    WScript.Echo ""
    
    WScript.Echo "MapPropKey:"
    WScript.Echo TestMapPropKey()
    
    WScript.Echo "NormalizeEnumToken:"
    WScript.Echo TestNormalizeEnumToken()
    
    WScript.Echo "ConvertColorToLong:"
    WScript.Echo TestConvertColorToLong()
    
    WScript.Echo "MapEventToProperty:"
    WScript.Echo TestMapEventToProperty()
End Sub

' Funciones implementadas directamente para prueba

Function MapPropKey(propName)
    Select Case LCase(propName)
        Case "ancho", "width": MapPropKey = "Width"
        Case "altura", "height": MapPropKey = "Height"
        Case "izquierda", "left": MapPropKey = "Left"
        Case "arriba", "top": MapPropKey = "Top"
        Case "nombre", "name": MapPropKey = "Name"
        Case "titulo", "caption": MapPropKey = "Caption"
        Case "visible": MapPropKey = "Visible"
        Case "habilitado", "enabled": MapPropKey = "Enabled"
        Case Else: MapPropKey = propName
    End Select
End Function

Function NormalizeEnumToken(token, propType)
    Select Case LCase(propType)
        Case "boolean"
            Select Case LCase(token)
                Case "verdadero", "true", "si", "yes", "1": NormalizeEnumToken = "True"
                Case "falso", "false", "no", "0": NormalizeEnumToken = "False"
                Case Else: NormalizeEnumToken = token
            End Select
        Case Else: NormalizeEnumToken = token
    End Select
End Function

Function ConvertColorToLong(colorHex)
    If Left(colorHex, 1) = "#" And Len(colorHex) = 7 Then
        Dim r, g, b
        r = CLng("&H" & Mid(colorHex, 2, 2))
        g = CLng("&H" & Mid(colorHex, 4, 2))
        b = CLng("&H" & Mid(colorHex, 6, 2))
        ConvertColorToLong = b * 65536 + g * 256 + r
    Else
        ConvertColorToLong = 0
    End If
End Function

Function MapEventToProperty(eventName)
    Select Case LCase(eventName)
        Case "clic", "click", "onclick": MapEventToProperty = "OnClick"
        Case "cargar", "load", "onload": MapEventToProperty = "OnLoad"
        Case "descargar", "unload", "onunload": MapEventToProperty = "OnUnload"
        Case "actual", "current", "oncurrent": MapEventToProperty = "OnCurrent"
        Case Else: MapEventToProperty = ""
    End Select
End Function

' Ejecutar pruebas
Main