
# Connection Close Method, Table Type Property Example (VB)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

Setting the [ActiveConnection](c1d90eca-9d62-4d7e-c275-5094e914ecb4.md) property to **Nothing** should "close" the catalog. Associated collections will be empty. Any objects that were created from schema objects in the catalog will be orphaned. Any properties on those objects that have been cached will still be available, but attempting to read properties that require a call to the provider will fail.




```vb
'BeginCloseConnectionVB 
Sub Main() 
 On Error GoTo CloseConnectionByNothingError 
 
 Dim cnn As New ADODB.Connection 
 Dim cat As New ADOX.Catalog 
 Dim tbl As ADOX.Table 
 
 cnn.Open "Provider='Microsoft.Jet.OLEDB.4.0';" &; _ 
 "Data Source= 'c:\Program Files\Microsoft Office\" &; _ 
 "Office\Samples\Northwind.mdb';" 
 Set cat.ActiveConnection = cnn 
 Set tbl = cat.Tables(0) 
 Debug.Print tbl.Type ' Cache tbl.Type info 
 Set cat.ActiveConnection = Nothing 
 Debug.Print tbl.Type ' tbl is orphaned 
 ' Previous line will succeed if this was cached 
 Debug.Print tbl.Columns(0).DefinedSize 
 ' Previous line will fail if this info has not been cached 
 
 'Clean up 
 cnn.Close 
 Set cat = Nothing 
 Set cnn = Nothing 
 Exit Sub 
 
CloseConnectionByNothingError: 
 Set cat = Nothing 
 
 If Not cnn Is Nothing Then 
 If cnn.State = adStateOpen Then cnn.Close 
 End If 
 Set cnn = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source &; "-->" &; Err.Description, , "Error" 
 End If 
End Sub 
' EndCloseConnectionVB 

```

Closing a [Connection](c16023aa-0321-2513-ee71-255d6ffba03d.md) object that was used to "open" the catalog should have the same effect as setting the **ActiveConnection** property to **Nothing**.



```vb
Sub CloseConnection() 
 On Error GoTo CloseConnectionError 
 
 Dim cnn As New ADODB.Connection 
 Dim cat As New ADOX.Catalog 
 Dim tbl As ADOX.Table 
 
 cnn.Open "Provider='Microsoft.Jet.OLEDB.4.0';" &; _ 
 "Data Source= 'c:\Program Files\Microsoft Office\" &; _ 
 "Office\Samples\Northwind.mdb';" 
 Set cat.ActiveConnection = cnn 
 Set tbl = cat.Tables(0) 
 Debug.Print tbl.Type ' Cache tbl.Type info 
 cnn.Close 
 Debug.Print tbl.Type ' tbl is orphaned 
 ' Previous line will succeed if this was cached 
 Debug.Print tbl.Columns(0).DefinedSize 
 ' Previous line will fail if this info has not been cached 
 
 'Clean up 
 Set cat = Nothing 
 Set cnn = Nothing 
 Exit Sub 
 
CloseConnectionError: 
 
 Set cat = Nothing 
 
 If Not cnn Is Nothing Then 
 If cnn.State = adStateOpen Then cnn.Close 
 End If 
 Set cnn = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source &; "-->" &; Err.Description, , "Error" 
 End If 
End Sub 
' EndCloseConnection2VB 

```

