
# Database.CreateTableDef Method (DAO)

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)


Creates a new  **[TableDef](715146b6-c62a-abff-28ee-e6bbe3c08adf.md)** object (Microsoft Access workspaces only). .

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **CreateTableDef**( ** _Name_**, ** _Attributes_**, ** _SourceTableName_**, ** _Connect_** )

 _expression_ A variable that represents a **Database** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**Variant**|A  **Variant** ( **String** subtype) that uniquely names the new **TableDef** object. See the **[Name](66b751ee-cf8a-a1f2-c646-6124e5f18cd0.md)** property for details on valid **TableDef** names.|
| _Attributes_|Optional|**Variant**|A constant or combination of constants that indicates one or more characteristics of the new  **TableDef** object. See the **[Attributes](d01588c3-e94e-06bd-6568-974873411f2d.md)** property for more information.|
| _SourceTableName_|Optional|**Variant**| A **Variant** ( **String** subtype) containing the name of a table in an external database that is the original source of the data. The _source_ string becomes the **[SourceTableName](3c02f5f6-70ae-39ec-0984-8d6b81992418.md)** property setting of the new **TableDef** object.|
| _Connect_|Optional|**Variant**|A  **Variant** ( **String** subtype) containing information about the source of an open database, a database used in a pass-through query, or a linked table. See the **[Connect](4fbb324c-a358-8fad-60f2-fb8005cf74d9.md)** property for more information about valid connection strings.|

### Return Value

TableDef


## Remarks
<a name="sectionSection1"> </a>

If you omit one or more of the optional parts when you use the  **CreateTableDef** method, you can use an appropriate assignment statement to set or reset the corresponding property before you append the new object to a collection. After you append the object, you can alter some but not all of its properties. See the individual property topics for more details.

If  _name_ refers to an object that is already a member of the collection, or you specify an invalid property in the **TableDef** or **[Field](47282ce2-9b49-ccf9-ad37-c4bb25cfd037.md)** object you're appending, a run-time error occurs when you use the **[Append](f951a3c4-dade-c1ef-3bfc-6b2a60e12adc.md)** method. Also, you can't append a **TableDef** object to the **TableDefs** collection until you define at least one **Field** for the **TableDef** object.

To remove a  **TableDef** object from the **[TableDefs](a2986b02-0437-d6ac-7bbb-c43f5225c3fc.md)** collection, use the **[Delete](130bb50d-17c3-b2ab-9360-0d91d0cee131.md)** method on the collection.


## Example
<a name="sectionSection2"> </a>

This example creates a new  **TableDef** object in the Northwind database.


```vb
Sub CreateTableDefX() 
 
 Dim dbsNorthwind As Database 
 Dim tdfNew As TableDef 
 Dim prpLoop As Property 
 
 Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 
 ' Create a new TableDef object. 
 Set tdfNew = dbsNorthwind.CreateTableDef("Contacts") 
 
 With tdfNew 
 ' Create fields and append them to the new TableDef 
 ' object. This must be done before appending the 
 ' TableDef object to the TableDefs collection of the 
 ' Northwind database. 
 .Fields.Append .CreateField("FirstName", dbText) 
 .Fields.Append .CreateField("LastName", dbText) 
 .Fields.Append .CreateField("Phone", dbText) 
 .Fields.Append .CreateField("Notes", dbMemo) 
 
 Debug.Print "Properties of new TableDef object " &; _ 
 "before appending to collection:" 
 
 ' Enumerate Properties collection of new TableDef 
 ' object. 
 For Each prpLoop In .Properties 
 On Error Resume Next 
 If prpLoop <> "" Then Debug.Print " " &; _ 
 prpLoop.Name &; " = " &; prpLoop 
 On Error GoTo 0 
 Next prpLoop 
 
 ' Append the new TableDef object to the Northwind 
 ' database. 
 dbsNorthwind.TableDefs.Append tdfNew 
 
 Debug.Print "Properties of new TableDef object " &; _ 
 "after appending to collection:" 
 
 ' Enumerate Properties collection of new TableDef 
 ' object. 
 For Each prpLoop In .Properties 
 On Error Resume Next 
 If prpLoop <> "" Then Debug.Print " " &; _ 
 prpLoop.Name &; " = " &; prpLoop 
 On Error GoTo 0 
 Next prpLoop 
 
 End With 
 
 ' Delete new TableDef object since this is a 
 ' demonstration. 
 dbsNorthwind.TableDefs.Delete "Contacts" 
 
 dbsNorthwind.Close 
 
End Sub
```

This example uses the  **CreateTableDef** and **FillCache** methods and the **CacheSize**, **CacheStart** and **SourceTableName** properties to enumerate the records in a linked table twice. Then it enumerates the records twice with a 50-record cache. The example then displays the performance statistics for the uncached and cached runs through the linked table.




```vb
Sub ClientServerX3() 
 
 Dim dbsCurrent As Database 
 Dim tdfRoyalties As TableDef 
 Dim rstRemote As Recordset 
 Dim sngStart As Single 
 Dim sngEnd As Single 
 Dim sngNoCache As Single 
 Dim sngCache As Single 
 Dim intLoop As Integer 
 Dim strTemp As String 
 Dim intRecords As Integer 
 
 ' Open a database to which a linked table can be 
 ' appended. 
 Set dbsCurrent = OpenDatabase("DB1.mdb") 
 
 ' Create a linked table that connects to a Microsoft SQL 
 ' Server database. 
 Set tdfRoyalties = _ 
 dbsCurrent.CreateTableDef("Royalties") 
 ' Note: The DSN referenced below must be set to 
 ' use Microsoft Windows NT Authentication Mode to 
 ' authorize user access to the Microsoft SQL Server. 
 tdfRoyalties.Connect = _ 
 "ODBC;DATABASE=pubs;DSN=Publishers" 
 tdfRoyalties.SourceTableName = "roysched" 
 dbsCurrent.TableDefs.Append tdfRoyalties 
 Set rstRemote = _ 
 dbsCurrent.OpenRecordset("Royalties") 
 
 With rstRemote 
 ' Enumerate the Recordset object twice and record 
 ' the elapsed time. 
 sngStart = Timer 
 
 For intLoop = 1 To 2 
 .MoveFirst 
 Do While Not .EOF 
 ' Execute a simple operation for the 
 ' performance test. 
 strTemp = !title_id 
 .MoveNext 
 Loop 
 Next intLoop 
 
 sngEnd = Timer 
 sngNoCache = sngEnd - sngStart 
 
 ' Cache the first 50 records. 
 .MoveFirst 
 .CacheSize = 50 
 .FillCache 
 sngStart = Timer 
 
 ' Enumerate the Recordset object twice and record 
 ' the elapsed time. 
 For intLoop = 1 To 2 
 intRecords = 0 
 .MoveFirst 
 Do While Not .EOF 
 ' Execute a simple operation for the 
 ' performance test. 
 strTemp = !title_id 
 ' Count the records. If the end of the 
 ' cache is reached, reset the cache to the 
 ' next 50 records. 
 intRecords = intRecords + 1 
 .MoveNext 
 If intRecords Mod 50 = 0 Then 
 .CacheStart = .Bookmark 
 .FillCache 
 End If 
 Loop 
 Next intLoop 
 
 sngEnd = Timer 
 sngCache = sngEnd - sngStart 
 
 ' Display performance results. 
 MsgBox "Caching Performance Results:" &; vbCr &; _ 
 " No cache: " &; Format(sngNoCache, _ 
 "##0.000") &; " seconds" &; vbCr &; _ 
 " 50-record cache: " &; Format(sngCache, _ 
 "##0.000") &; " seconds" 
 .Close 
 End With 
 
 ' Delete linked table because this is a demonstration. 
 dbsCurrent.TableDefs.Delete tdfRoyalties.Name 
 dbsCurrent.Close 
 
End Sub
```

