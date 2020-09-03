Attribute VB_Name = "mod_ProcesaBD"
Option Explicit

Public Function TraerPais(ByVal strCodPais As String, Optional ByRef o_lError As Long) As ADODB.Recordset
       
       Dim rsAux As ADODB.Recordset
       Dim cmdSQL As New ADODB.Command
       
1000   On Error GoTo ControlError
       
1010   Set rsAux = New ADODB.Recordset
1020   rsAux.CursorLocation = adUseClient 'soy cliente de la bd
1030   With cmdSQL
1040      .ActiveConnection = ConexSQL
1050      .CommandType = adCmdStoredProc
1060      .CommandText = "sp_buscar_pais"
1070      .Parameters.Append .CreateParameter("@cod_pais", adVarChar, adParamInput, 10, strCodPais)
1080      Set rsAux = cmdSQL.Execute
1090      .ActiveConnection = Nothing
1100   End With
1110   Set TraerPais = rsAux
1120   Exit Function
ControlError:
1130   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
End Function

Public Function TraerDepto(ByVal strCodPais As String, ByVal strCodDepto As String, Optional ByRef o_lError As Long) As ADODB.Recordset
       
       Dim rsAux As ADODB.Recordset
       Dim cmdSQL As New ADODB.Command
       
1000   On Error GoTo ControlError
       
1010   Set rsAux = New ADODB.Recordset
1020   rsAux.CursorLocation = adUseClient 'soy cliente de la bd
1030   With cmdSQL
1040      .ActiveConnection = ConexSQL
1050      .CommandType = adCmdStoredProc
1060      .CommandText = "sp_buscar_depto"
1070      .Parameters.Append .CreateParameter("@codPais", adVarChar, adParamInput, 10, strCodPais)
1080      .Parameters.Append .CreateParameter("@codDepto", adVarChar, adParamInput, 10, strCodDepto)
1090      Set rsAux = cmdSQL.Execute
1100      .ActiveConnection = Nothing
1110   End With
1120   Set TraerDepto = rsAux
1130   Exit Function
ControlError:
1140   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
End Function

Public Function TraerDeptoDesc(ByVal strCodPais As String, ByVal strDescDepto As String, Optional ByRef o_lError As Long) As ADODB.Recordset
       
       Dim rsAux As ADODB.Recordset
       Dim cmdSQL As New ADODB.Command
       
1000   On Error GoTo ControlError
       
1010   Set rsAux = New ADODB.Recordset
1020   rsAux.CursorLocation = adUseClient 'soy cliente de la bd
1030   With cmdSQL
1040      .ActiveConnection = ConexSQL
1050      .CommandType = adCmdStoredProc
1060      .CommandText = "sp_buscar_depto_desc"
1070      .Parameters.Append .CreateParameter("@codPais", adVarChar, adParamInput, 20, strCodPais)
1080      .Parameters.Append .CreateParameter("@desDepto", adVarChar, adParamInput, 150, strDescDepto)
1090      Set rsAux = cmdSQL.Execute
1100      .ActiveConnection = Nothing
1110   End With
1120   Set TraerDeptoDesc = rsAux
1130   Exit Function
ControlError:
1140   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
End Function

Public Function TraerPaisDesc(ByVal strDescPais As String, Optional ByRef o_lError As Long) As ADODB.Recordset
       
       Dim rsAux As ADODB.Recordset
       Dim cmdSQL As New ADODB.Command
       
1000   On Error GoTo ControlError
       
1010   Set rsAux = New ADODB.Recordset
1020   rsAux.CursorLocation = adUseClient 'soy cliente de la bd
1030   With cmdSQL
1040      .ActiveConnection = ConexSQL
1050      .CommandType = adCmdStoredProc
1060      .CommandText = "sp_buscar_pais_desc"
1070      .Parameters.Append .CreateParameter("@desc_pais", adVarChar, adParamInput, 120, strDescPais)
1080      Set rsAux = cmdSQL.Execute
1090      .ActiveConnection = Nothing
1100   End With
1110   Set TraerPaisDesc = rsAux
1120   Exit Function
ControlError:
1130   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
End Function

