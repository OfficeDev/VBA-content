
# Making a Connection

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

To connect to a data source, you must specify a  _connection string_, the parameters of which might differ for each provider and data source. For more information, see[Creating the Connection String](0d34b1c6-bf2e-1299-9778-573ccd2da1c7.md).

ADO most commonly opens a connection by using the  **Connection** object **Open** method. The syntax for the **Open** method is shown here:



```
 
Dim connection as New ADODB.Connection 
connection.OpenConnectionString , UserID , Password , OpenOptions
```

Alternatively, you can invoke a shortcut technique,  **Recordset.Open**, to open an implicit connection and issue a command over that connection in one operation. Do this by passing in a valid connection string as the _ActiveConnection_ argument to the **Open** method. Here is the syntax for each method in Visual Basic:



```vb
 
Dim recordset as ADODB.Recordset 
Set recordset = New ADODB.Recordset 
recordset.OpenSource , ActiveConnection , CursorType , LockType , Options
```


 **Note**  When should you use a  **Connection** object vs. the **Recordset.Open** shortcut? Use the **Connection** object if you plan to open more than one **Recordset**, or when executing multiple commands. A connection is still created by ADO implicitly when you use the **Recordset.Open** shortcut.

