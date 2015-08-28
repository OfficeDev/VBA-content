
# Bind a Form to an ADO Recordset

 **Last modified:** July 30, 2015

 _**Applies to:** Access 2013_

To bind an Access form to a recordset, you must set the form's  **Recordset** property to an open ADO **Recordset** object. A form must meet two general requirements for the form to be updatable when it is bound to an ADO recordset. The general requirements are:


- The underlying ADO recordset must be updatable via ADO.
    
- The recordset must contain one or more fields that are uniquely indexed, such as a table's primary key.
    




```
 Private Sub Form_Open(Cancel As Integer) 
 Dim cn As ADODB.Connection 
 Dim rs As ADODB.Recordset 
 
 'Use the ADO connection that Access uses 
 Set cn = CurrentProject.AccessConnection 
 'Create an instance of the ADO Recordset class, 
 'and set its properties 
 Set rs = New ADODB.Recordset 
 With rs 
 Set .ActiveConnection = cn 
 .Source = "SELECT * FROM Customers" 
 .LockType = adLockOptimistic 
 .CursorType = adOpenKeyset 
 .Open 
 End With 
 'Set the form's Recordset property to the ADO recordset 
 Set Me.Recordset = rs 
 Set rs = Nothing 
 Set cn = Nothing 
 End Sub 

```

