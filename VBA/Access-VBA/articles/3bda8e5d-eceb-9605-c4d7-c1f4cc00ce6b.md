
# The Fields Collection

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Examining the Fields Collection](#sectionSection0)
[Counting Columns](#sectionSection1)
[Getting to the Field](#sectionSection2)
[Using the Refresh Method](#sectionSection3)
[Adding Fields to a Recordset](#sectionSection4)


The  **Fields** collection is one of ADO's intrinsic collections. A collection is an ordered set of items that can be referred to as a unit.
The  **Fields** collection contains a **Field** object for every field (column) in the **Recordset**. Like all ADO collections, it has **Count** and **Item** properties, as well as **Append** and **Refresh** methods. It also has **CancelUpdate**, **Delete**, **Resync**, and **Update** methods, which are not available to other ADO collections.

## Examining the Fields Collection
<a name="sectionSection0"> </a>

Consider the  **Fields** collection of the sample **Recordset** introduced in this chapter. The sample **Recordset** was derived from the SQL statement


```sql
SELECT ProductID, ProductName, UnitPrice FROM Products WHERE CategoryID = 7 
```

Thus, you should find that the  **Recordset** **Fields** collection contains three fields.




```vb
'BeginWalkFields 
 Dim objFields As ADODB.Fields 
 
 objRs.Open strSQL, strConnStr, adOpenForwardOnly, adLockReadOnly, adCmdText 
 
 Set objFields = objRs.Fields 
 
 For intLoop = 0 To (objFields.Count - 1) 
 Debug.Print objFields.Item(intLoop).Name 
 Next 
'EndWalkFields 

```

This code simply determines the number of  **Field** objects in the **Fields** collection using the **Count** property and loops through the collection, returning the value of the **Name** property for each **Field** object. You can use many more **Field** properties to get information about a field. For more information about querying a **Field**, see[The Field Object](55531e04-d74f-6394-df64-1660e5d572ca.md).


## Counting Columns
<a name="sectionSection1"> </a>

As you might expect, the  **Count** property returns the actual number of **Field** objects in the **Fields** collection. Because numbering for members of a collection begins with zero, you should always code loops starting with the zero member and ending with the value of the **Count** property minus 1. If you are using Microsoft Visual Basic and want to loop through the members of a collection without checking the **Count** property, use the **For** **Each...Next** command.

If the  **Count** property is zero, there are no objects in the collection.


## Getting to the Field
<a name="sectionSection2"> </a>

As with any ADO collection, the  **Item** property is the default property of the collection. It returns the individual **Field** object specified by the name or index passed to it. Therefore, the following statements are equivalent for the sample **Recordset**:


```
 
objField = objRecordset.Fields.Item("ProductID") 
objField = objRecordset.Fields("ProductID") 
objField = objRecordset.Fields.Item(0) 
objField = objRecordset.Fields(0) 

```

If these methods are equivalent, which is best? It depends. Using an index to retrieve a  **Field** from the collection is faster because it accesses the **Field** directly without having to perform a string lookup. On the other hand, the order of **Fields** within the collection must be known, and if the order changes, the reference to the **Field's** index will have to be changed wherever it occurs. Although slightly slower, using the name of the **Field** is more flexible because it doesn't depend on the order of the **Fields** in the collection.


## Using the Refresh Method
<a name="sectionSection3"> </a>

Unlike some other ADO collections, using the  **Refresh** method on the **Fields** collection has no visible effect. To retrieve changes from the underlying database structure, you must use either the **Requery** method, or if the **Recordset** object does not support bookmarks, the **MoveFirst** method, which will cause the command to be executed against the provider again.


## Adding Fields to a Recordset
<a name="sectionSection4"> </a>

The  **Append** method is used to add fields to a **Recordset**.

You can use the  **Append** method to fabricate a **Recordset** programmatically without opening a connection to a data source. A run-time error will occur if the **Append** method is called on the **Fields** collection of an open **Recordset** or on a **Recordset** where the **ActiveConnection** property has been set. You can append fields only to a **Recordset** that is not open and has not yet been connected to a data source. However, to specify values for the newly appended **Fields**, the **Recordset** must first be opened.

Developers often need a place to temporarily store some data, or want some data to act as if it came from a server so it can participate in data binding in a user interface. ADO (in conjunction with the [Microsoft Cursor Service for OLE DB](6818fc05-9c9f-9b67-07d2-e622c93133c2.md)) enables the developer to build an empty  **Recordset** object by specifying column information and calling **Open**. In the following example, three new fields are appended to a new **Recordset** object. Then the **Recordset** is opened, two new records are added, and the **Recordset** is persisted to a file. (For more information about **Recordset** persistence, see[Chapter 5: Updating and Persisting Data](77acb763-1c60-1945-791d-3e83d684fb0d.md).)




```vb
'BeginFabricate 
 Dim objRs As New ADODB.Recordset 
 
 With objRs.Fields 
 .Append "StudentID", adChar, 11, adFldUpdatable 
 .Append "FullName", adVarChar, 50, adFldUpdatable 
 .Append "PhoneNmbr", adVarChar, 20, adFldUpdatable 
 End With 
 
 With objRs 
 .Open 
 
 .AddNew 
 .Fields(0) = "123-45-6789" 
 .Fields(1) = "John Doe" 
 .Fields(2) = "(425) 555-5555" 
 .Update 
 
 .AddNew 
 .Fields(0) = "123-45-6780" 
 .Fields(1) = "Jane Doe" 
 .Fields(2) = "(615) 555-1212" 
 .Update 
 End With 
 
 objRs.Save App.Path &; "\FabriTest.adtg", adPersistADTG 
 
 objRs.Close 
 'EndFabricate 

```

The usage of the  **Fields** **Append** method differs between the **Recordset** object and the **Record** object. For more information about the **Record** object, see[Chapter 10: Records and Streams](74862096-2273-3b61-f89c-06554ccf42cd.md).

