Attribute VB_Name = "mod_ProcesaBD"
Option Explicit

Public Function TraerPais(ByVal strCodPais As String, Optional ByRef o_lError As Long) As ADODB.Recordset

    Dim rsAux As ADODB.Recordset
    Dim cmdSQL As New ADODB.Command
       
    On Error GoTo ControlError
       
    Set rsAux = New ADODB.Recordset
    rsAux.CursorLocation = adUseClient 'soy cliente de la bd
        With cmdSQL
            .ActiveConnection = ConexSQL
            .CommandType = adCmdStoredProc
            .CommandText = "sp_buscar_pais"
            .Parameters.Append .CreateParameter("@cod_pais", adVarChar, adParamInput, 10, strCodPais)
            Set rsAux = cmdSQL.Execute
            .ActiveConnection = Nothing
        End With
    Set TraerPais = rsAux
    Exit Function
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
End Function

Public Function TraerDepto(ByVal strCodPais As String, ByVal strCodDepto As String, Optional ByRef o_lError As Long) As ADODB.Recordset

    Dim rsAux As ADODB.Recordset
    Dim cmdSQL As New ADODB.Command
       
    On Error GoTo ControlError
       
    Set rsAux = New ADODB.Recordset
    rsAux.CursorLocation = adUseClient 'soy cliente de la bd
        With cmdSQL
            .ActiveConnection = ConexSQL
            .CommandType = adCmdStoredProc
            .CommandText = "sp_buscar_depto"
            .Parameters.Append .CreateParameter("@codPais", adVarChar, adParamInput, 10, strCodPais)
            .Parameters.Append .CreateParameter("@codDepto", adVarChar, adParamInput, 10, strCodDepto)
            Set rsAux = cmdSQL.Execute
            .ActiveConnection = Nothing
        End With
    Set TraerDepto = rsAux
    Exit Function
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
End Function

Public Function TraerDeptoDesc(ByVal strCodPais As String, ByVal strDescDepto As String, Optional ByRef o_lError As Long) As ADODB.Recordset

    Dim rsAux As ADODB.Recordset
    Dim cmdSQL As New ADODB.Command
       
    On Error GoTo ControlError
       
    Set rsAux = New ADODB.Recordset
    rsAux.CursorLocation = adUseClient 'soy cliente de la bd
        With cmdSQL
            .ActiveConnection = ConexSQL
            .CommandType = adCmdStoredProc
            .CommandText = "sp_buscar_depto_desc"
            .Parameters.Append .CreateParameter("@codPais", adVarChar, adParamInput, 20, strCodPais)
            .Parameters.Append .CreateParameter("@desDepto", adVarChar, adParamInput, 150, strDescDepto)
            Set rsAux = cmdSQL.Execute
            .ActiveConnection = Nothing
        End With
    Set TraerDeptoDesc = rsAux
    Exit Function
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
End Function

Public Function TraerPaisDesc(ByVal strDescPais As String, Optional ByRef o_lError As Long) As ADODB.Recordset

    Dim rsAux As ADODB.Recordset
    Dim cmdSQL As New ADODB.Command
       
    On Error GoTo ControlError
       
    Set rsAux = New ADODB.Recordset
    rsAux.CursorLocation = adUseClient 'soy cliente de la bd
        With cmdSQL
            .ActiveConnection = ConexSQL
            .CommandType = adCmdStoredProc
            .CommandText = "sp_buscar_pais_desc"
            .Parameters.Append .CreateParameter("@desc_pais", adVarChar, adParamInput, 120, strDescPais)
            Set rsAux = cmdSQL.Execute
            .ActiveConnection = Nothing
        End With
    Set TraerPaisDesc = rsAux
    Exit Function
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
End Function

