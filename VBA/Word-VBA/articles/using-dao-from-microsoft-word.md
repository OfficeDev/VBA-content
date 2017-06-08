---
title: Using DAO from Microsoft Word
ms.prod: word
ms.assetid: f8c2b535-b912-e7ff-73a0-3b6558aae565
ms.date: 06/08/2017
---


# Using DAO from Microsoft Word

You can use Data Access Objects (DAO) properties, objects, and methods the same way that you reference and use Word properties, objects, and methods. After you establish a reference to the DAO object library, you can open databases, design and run queries to extract a set of records, and bring the results back to Word.


## Referencing DAO

Before you can use DAO, you must establish a reference to the DAO object library. Use the following steps to establish a reference to the DAO object library.


1. Switch to the Visual Basic Editor.
    
2. On the  **Tools** menu, click **References**.
    
3. In the  **Available References** box, select **Microsoft DAO 3.6 Object Library**.
    
The following example opens the Northwind database and inserts the items from the Shippers table into the active document.




```vb
Sub UsingDAOWithWord() 
 Dim docNew As Document 
 Dim dbNorthwind As DAO.Database 
 Dim rdShippers As Recordset 
 Dim intRecords As Integer 
 
 Set docNew = Documents.Add 
 Set dbNorthwind = OpenDatabase _ 
 (Name:="C:\Program Files\Microsoft Office\Office11\" _ 
 &; "Samples\Northwind.mdb") 
 Set rdShippers = dbNorthwind.OpenRecordset(Name:="Shippers") 
 For intRecords = 0 To rdShippers.RecordCount - 1 
 docNew.Content.InsertAfter Text:=rdShippers.Fields(1).Value 
 rdShippers.MoveNext 
 docNew.Content.InsertParagraphAfter 
 Next intRecords 
 rdShippers.Close 
 dbNorthwind.Close 
End Sub
```

Use the  **OpenDatabase** method to connect to a database and open it. After opening the database, use the **OpenRecordset** method to access a table or query for results. To navigate through the recordset, use the **Move** method. To find a specific record, use the **Seek** method. If you need only a subset of records instead of the entire recordset, use the **CreateQueryDef** method to design a customized query to select records that meet your criteria. When you finish working with a database, it is a good idea to close it using the **Close** method, to save memory.


## Remarks

For more information about a specific DAO object, method, or property, see the information about Data Access Objects on MSDN.


