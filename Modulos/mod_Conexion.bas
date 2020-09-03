Attribute VB_Name = "mod_Conexion"
Sub main()
1000   On Error GoTo ControlError
       'Conectar a la BD
1010   With ConexSQL
1020      .CursorLocation = adUseClient 'Soy cliente de la BD
1030      .Open "Provider= SQLOLEDB.1;" & _
             "Integrated Security= SSPI;" & _
             "Persist Security Info= false;" & _
             "Initial Catalog= almCarros;" & _
             "Data Source= ASUSK555D\SQLEXPRESS;"
          '      "Data Source= CLIENTE-PC;"
          'ASUSK555D\SQLEXPRESS
          'frmPais.Show
          'frmDepto.Show
1040      frmCiudad.Show
          
1050   End With
ExitProc:
1060   Exit Sub
ControlError:
1070   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1080   Resume ExitProc
End Sub

'Provider=SQLOLEDB.1;Password=1030538949;Persist Security Info=True;User ID=ccgutierrezm;Initial Catalog=almCarros;Data Source=CLIENTE-PC
'Provider=MSDASQL.1;Password=1030538949;Persist Security Info=True;User ID=ccgutierrezm;Data Source=SQL Server;Initial Catalog=almCarros
