
# Persisting Data

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

Portable computing (for example, using laptops) has generated the need for applications that can run in both a connected and disconnected state. ADO has added support for this by giving the developer the ability to save a client cursor  **Recordset** to disk and reload it later.

There are several scenarios in which you could use this type of feature, including the following:

-  **Traveling:** When taking the application on the road, it is vital to supply the ability to make changes and add new records that can then be reconnected to the database later and committed.
    
-  **Infrequently updated lookups:** Often in an application, tables are used as lookups — for example, state tax tables. They are infrequently updated and are read-only. Instead of rereading this data from the server each time the application is started, the application can simply load the data from a locally persisted **Recordset**.
    
In ADO, to save and load  **Recordsets**, use the **Recordset.Save** and **Recordset.Open(,,,,adCmdFile)** methods on the ADO **Recordset** object.
You can use the  **Recordset** **Save** method to persist your ADO **Recordset** to a file on a disk. (You can also save a **Recordset** to an ADO **Stream** object. **Stream** objects are discussed later in the guide.) Later, you can use the **Open** method to reopen the **Recordset** when you are ready to use it. By default, ADO saves the **Recordset** into the proprietary Microsoft Advanced Data TableGram (ADTG) format. This binary format is specified using the **adPersistADTG** **PersistFormatEnum** value. Alternatively, you may choose to save your **Recordset** out as XML instead using **adPersistXML**. For more information about saving Recordsets as XML, see[Persisting Records in XML Format](8071e244-60c7-759c-094c-152add5d72e4.md).
The syntax of the  **Save** method is as follows:



```
recordset.SaveDestination, PersistFormat
```

The first time you save the  **Recordset**, it is optional to specify _Destination_. If you omit _Destination_, a new file will be created with a name set to the value of the[Source](523ea81e-d011-8d87-436e-084b6eba0908.md) property of the **Recordset**.
Omit  _Destination_ when you subsequently call **Save** after the first save or a run-time error will occur. If you subsequently call **Save** with a new _Destination_, the **Recordset** is saved to the new destination. However, the new destination and the original destination will both be open.
 **Save** does not close the **Recordset** or _Destination_, so you can continue to work with the **Recordset** and save your most recent changes. _Destination_ remains open until the **Recordset** is closed, during which time other applications can read but not write to _Destination_.
For reasons of security, the  **Save** method permits only the use of low and custom security settings from a script executed by Microsoft Internet Explorer. For a more detailed explanation of security issues, see "ADO and RDS Security Issues in Microsoft Internet Explorer" under ActiveX Data Objects (ADO) Technical Articles in Microsoft Data Access Technical Articles.
If the  **Save** method is called while an asynchronous **Recordset** fetch, execute, or update operation is in progress, **Save** waits until the asynchronous operation is complete.
Records are saved beginning with the first row of the  **Recordset**. When the **Save** method is finished, the current row position is moved to the first row of the **Recordset**.
For best results, set the [CursorLocation](8a048bd4-ae25-a555-1c07-14364b7e6560.md) property to **adUseClient** with **Save**. If your provider does not support all of the functionality necessary to save **Recordset** objects, the Cursor Service will provide that functionality.
When a  **Recordset** is persisted with the **CursorLocation** property set to **adUseServer**, the update capability for the **Recordset** is limited. Typically, only single-table updates, insertions, and deletions are allowed (dependent on provider functionality). The[Resync](f594a200-56e6-fcf5-9b0a-900c56377f24.md) method is also unavailable in this configuration.
Because the  _Destination_ parameter can accept any object that supports the OLE DB **IStream** interface, you can save a **Recordset** directly to the ASP **Response** object.
In the following example, the  **Save** and **Open** methods are used to persist a **Recordset** and later reopen it:



```vb
'BeginPersist 
 conn.ConnectionString = _ 
 "Provider='SQLOLEDB';Data Source='MySqlServer';" _ 
 &; "Integrated Security='SSPI';Initial Catalog='pubs'" 
 conn.Open 
 
 conn.Execute "create table testtable (dbkey int " &; _ 
 "primary key, field1 char(10))" 
 conn.Execute "insert into testtable values (1, 'string1')" 
 
 Set rst.ActiveConnection = conn 
 rst.CursorLocation = adUseClient 
 
 rst.Open "select * from testtable", conn, adOpenStatic, _ 
 adLockBatchOptimistic 
 
 'Change the row on the client 
 rst!field1 = "NewValue" 
 
 'Save to a file--the .dat extension is an example; choose 
 'your own extension. The changes will be saved in the file 
 'as well as the original data. 
 MyFile = Dir("c:\temp\temptbl.dat") 
 If MyFile <> "" Then 
 Kill "c:\temp\temptbl.dat" 
 End If 
 
 rst.Save "c:\temp\temptbl.dat", adPersistADTG 
 rst.Close 
 Set rst = Nothing 
 
 'Now reload the data from the file 
 Set rst = New ADODB.Recordset 
 rst.Open "c:\temp\temptbl.dat", , adOpenStatic, _ 
 adLockBatchOptimistic, adCmdFile 
 
 Debug.Print "After Loading the file from disk" 
 Debug.Print " Current Edited Value: " &; rst!field1.Value 
 Debug.Print " Value Before Editing: " &; rst!field1.OriginalValue 
 
 'Note that you can reconnect to a connection and 
 'submit the changes to the data source 
 Set rst.ActiveConnection = conn 
 rst.UpdateBatch 
'EndPersist 

```

