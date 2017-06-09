---
title: AllStoredProcedures Object (Access)
keywords: vbaac10.chm12691
f1_keywords:
- vbaac10.chm12691
ms.prod: access
api_name:
- Access.AllStoredProcedures
ms.assetid: 896f4c2c-273c-2849-0f06-d75fa515c44a
ms.date: 06/08/2017
---


# AllStoredProcedures Object (Access)

The  **AllStoredProcedures** collection contains an **[AccessObject](accessobject-object-access.md)** for each stored procedure in the **[CurrentData](currentdata-object-access.md)** or **[CodeData](codedata-object-access.md)** object.


## Remarks

The  **CurrentData** or **CodeData** object has an **AllStoredProcedures** collection containing **AccessObject** objects that describe instances of all stored procedures specified by **CurrentData** or **CodeData**. For example, you can enumerate the **AllStoredProcedures** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual  **AccessObject** object in the **AllStoredProcedures** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllStoredProcedures** collection, it's better to refer to the stored procedures by name because a stored procedure's collection index may change.

The  **AllStoredProcedures** collection is indexed beginning with zero. If you refer to a stored procedure by its index, the first stored procedure is AllStoredProcedures(0), the second stored procedure is AllStoredProcedures(1), and so on.


 **Note**  


## Example

The following example prints the name of each open  **AccessObject** object in the **AllProcedures** collection.


```
Sub AllStoredProcedures() 
    Dim obj As AccessObject, dbs As Object 
    Set dbs = Application.CurrentData 
    ' Search for open AccessObject objects in 
    ' AllStoredProcedures collection. 
    For Each obj In dbs.AllStoredProcedures 
        If obj.IsLoaded = True Then 
            ' Print name of obj. 
            Debug.Print obj.Name 
        End If 
    Next obj 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](allstoredprocedures-application-property-access.md)|
|[Count](allstoredprocedures-count-property-access.md)|
|[Item](allstoredprocedures-item-property-access.md)|
|[Parent](allstoredprocedures-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
