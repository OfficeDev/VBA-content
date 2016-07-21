
# Alternatives: Using SQL Statements

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

ADO also allows using commands as alternatives to its built-in properties and methods for editing data. Depending upon your provider, all operations mentioned in this chapter could also be accomplished by passing commands to your data source. For example, SQL UPDATE statements can be used to modify data without using the  **Value** property of a **Field**. SQL INSERT statements can be used to add new records to a data source, rather than the ADO method **AddNew**. For more information about SQL or the data-manipulation language of your provider, see the documentation of your data source.

For example, you can pass a SQL string containing a DELETE statement to a database, as shown in the following code:



```vb
'BeginSQLDelete 
strSQL = "DELETE FROM Shippers WHERE ShipperID = " &; intId 
objConn.Execute strSQL, , adCmdText + adExecuteNoRecords 
'EndSQLDelete 

```

