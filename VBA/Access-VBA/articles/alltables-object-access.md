---
title: AllTables Object (Access)
keywords: vbaac10.chm12688
f1_keywords:
- vbaac10.chm12688
ms.prod: access
api_name:
- Access.AllTables
ms.assetid: 530bff2d-1d0b-4790-a0f4-ffc628e7f130
ms.date: 06/08/2017
---


# AllTables Object (Access)

The  **AllTables** collection contains an **[AccessObject](accessobject-object-access.md)** for each table in the **[CurrentData](http://msdn.microsoft.com/library/c8d3f74f-050d-e1be-9496-2f1e20996066%28Office.15%29.aspx)** or **[CodeData](http://msdn.microsoft.com/library/fc207136-4d18-2c7d-ffe6-0e1ad7c2fc32%28Office.15%29.aspx)** object.


## Remarks

The  **CurrentData** or **CodeData** object has an **AllTables** collection containing **AccessObject** objects that describe instances of all tables specified by **CurrentData** or **CodeData**. For example, you can enumerate the **AllTables** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual  **AccessObject** object in the **AllTables** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllTables** collection, it's better to refer to the table by name because a table's collection index may change.

The  **AllTables** collection is indexed beginning with zero. If you refer to a table by its index, the first table is AllTables(0), the second table is AllTables(1), and so on.


 **Note**   To list all open tables in the database, use the **[IsLoaded](http://msdn.microsoft.com/library/5e68398c-8a95-f3e1-87ec-e2d637f34429%28Office.15%29.aspx)** property of each **AccessObject** object in the **AllTables** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a table.

You can't add or delete an  **AccessObject** object from the **AllTables** collection.


## Example

The following example prints the name of each open  **AccessObject** object in the **AllTables** collection.


```
Sub AllTables() 
 Dim obj As AccessObject, dbs As Object 
 Set dbs = Application.CurrentData 
 ' Search for open AccessObject objects in AllTables collection. 
 For Each obj In dbs.AllTables 
 If obj.IsLoaded = True Then 
 ' Print name of obj. 
 Debug.Print obj.Name 
 End If 
 Next obj 
End Sub
```



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/2c5c5433-abca-1c36-af82-3b5dbf12c793%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/f89a096e-e88c-ffcf-430b-8736c60d5bb9%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/87e8e90d-abfe-c15b-f450-c82a8292fc93%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/41801b04-da13-c917-8692-d6d56bd45516%28Office.15%29.aspx)|

## See also


#### Other resources


[AllTables Object Members](http://msdn.microsoft.com/library/29ac5838-ff13-b187-8f1e-54e7a533d084%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
