
# Counting Rows

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

The  **RecordCount** property returns a **Long** value that indicates the number of records in the **Recordset**. Use the **RecordCount** property to find out how many records are in a **Recordset** object. The property returns -1 when ADO cannot determine the number of records or if the provider or cursor type does not support **RecordCount**. Reading the **RecordCount** property on a closed **Recordset** causes an error.

The  **RecordCount** property depends on the capabilities of the provider and the type of cursor. The **RecordCount** property will return -1 for a forward-only cursor, the actual count for a static or keyset cursor, and either -1 or the actual count for a dynamic cursor, depending on the data source.
The sample  **Recordset** introduced in[Examining Data](73c69134-3127-3344-d5c3-5ecb9e0e958b.md) would return -1 because a forward-only cursor was opened. In order to use the **RecordCount** property, you would need to open the **Recordset** with a more sophisticated cursor (static or keyset).
In certain cases, your provider or cursor might be unable to provide the  **RecordCount** value without first fetching all records from the data source. To force this type of fetch, call the **Recordset** **MoveLast** method before calling **RecordCount**.
If you were to replace the line of code that calls the  **Recordset** **Open** method with the following:



```
 
oRs.Open sSQL, sCnStr, adOpenStatic, adLockOptimistic, adCmdText 

```

you would be able to use the  **RecordCount** property because static cursors with the[Microsoft OLE DB Provider for SQL Server](0ffdea03-1a76-499b-f649-423f6b3c13d7.md) support **RecordCount**. For example, the following code would print out the number of records returned by the command to the debug window, assuming the cursor supports the **RecordCount** property:



```
 
Debug.Print oRs.RecordCount ' Output: 4 

```

From this point forward, assume that these more capable (but more expensive) cursor and lock type settings are used.
