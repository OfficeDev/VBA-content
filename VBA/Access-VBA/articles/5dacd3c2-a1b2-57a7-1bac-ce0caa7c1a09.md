
# Views Collection, CommandText Property Example (VB)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

The following code demonstrates how to use the [Command](96b5fc97-f426-8540-fce0-2319ac8524a5.md) property to update the text of a view.




```vb
'BeginViewsCollectionVB 
Sub Main() 
 On Error GoTo ViewTextError 
 
 Dim cnn As New ADODB.Connection 
 Dim cat As New ADOX.Catalog 
 Dim cmd As New ADODB.Command 
 
 ' Open the Connection 
 cnn.Open _ 
 "Provider='Microsoft.Jet.OLEDB.4.0';" &; _ 
 "Data Source='c:\Program Files\Microsoft Office\" &; _ 
 "Office\Samples\Northwind.mdb';" 
 
 ' Open the catalog 
 Set cat.ActiveConnection = cnn 
 
 ' Get the command 
 Set cmd = cat.Views("AllCustomers").Command 
 
 ' Update the CommandText of the Command 
 cmd.CommandText = _ 
 "Select CustomerId, CompanyName, ContactName From Customers" 
 
 ' Update the View 
 Set cat.Views("AllCustomers").Command = cmd 
 
 'Clean up 
 cnn.Close 
 Set cat = Nothing 
 Set cmd = Nothing 
 Set cnn = Nothing 
 Exit Sub 
 
ViewTextError: 
 
 Set cat = Nothing 
 Set cmd = Nothing 
 
 If Not cnn Is Nothing Then 
 If cnn.State = adStateOpen Then cnn.Close 
 End If 
 Set cnn = Nothing 
 
 If Err <> 0 Then 
 MsgBox Err.Source &; "-->" &; Err.Description, , "Error" 
 End If 
End Sub 
' EndViewsCollectionVB 

```

