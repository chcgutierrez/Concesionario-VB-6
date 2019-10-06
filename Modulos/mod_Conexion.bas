Attribute VB_Name = "mod_Conexion"
Sub main()
On Error GoTo ControlError
'Conectar a la BD
With ConexSQL
.CursorLocation = adUseClient 'Soy cliente de la BD
.Open "Provider= SQLOLEDB.1;" & _
      "Integrated Security= SSPI;" & _
      "Persist Security Info= false;" & _
      "Initial Catalog= almCarros;" & _
      "Data Source= ASUSK555D\SQLEXPRESS;"
'      "Data Source= CLIENTE-PC;"
'ASUSK555D\SQLEXPRESS
frmDepto.Show
'frmPais.Show
End With
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

'Provider=SQLOLEDB.1;Password=1030538949;Persist Security Info=True;User ID=ccgutierrezm;Initial Catalog=almCarros;Data Source=CLIENTE-PC
'Provider=MSDASQL.1;Password=1030538949;Persist Security Info=True;User ID=ccgutierrezm;Data Source=SQL Server;Initial Catalog=almCarros
