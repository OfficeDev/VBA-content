---
title: AllDatabaseDiagrams Object (Access)
keywords: vbaac10.chm12692
f1_keywords:
- vbaac10.chm12692
ms.prod: access
api_name:
- Access.AllDatabaseDiagrams
ms.assetid: 417427aa-1783-29da-30c9-66a7032a0088
ms.date: 06/08/2017
---


# AllDatabaseDiagrams Object (Access)

The  **AllDatabaseDiagrams** collection contains an **[AccessObject](accessobject-object-access.md)** for each database diagram in the **[CurrentData](currentdata-object-access.md)** or **[CodeData](codedata-object-access.md)** object.


## Remarks

The  **CurrentData** or **CodeData** object has an **AllDatabaseDiagrams** collection containing **AccessObject** objects that describe instances of all database diagrams specified by **CurrentData** or **CodeData**. For example, you can enumerate the **AllDatabaseDiagrams** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual  **AccessObject** object in the **AllDatabaseDiagrams** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllDatabaseDiagrams** collection, it's better to refer to the database diagram by name because a database diagram's collection index may change.

The  **AllDatabaseDiagrams** collection is indexed beginning with zero. If you refer to a database diagram by its index, the first database diagram is AllDatabaseDiagrams(0), the second database diagram is AllDatabaseDiagrams(1), and so on.


 **Note**  

You can't add or delete an  **AccessObject** object from the **AllDatabaseDiagrams** collection.


## Example

The following example prints the name of each open  **AccessObject** object in the **AllDatabaseDiagrams** collection.


```
Sub AllDatabaseDiagrams() 
    Dim obj As AccessObject, dbs As Object 
    Set dbs = Application.CurrentData 
    ' Search for open AccessObject objects in 
    ' AllDatabaseDiagrams collection. 
    For Each obj In dbs.AllDatabaseDiagrams 
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
|[Application](alldatabasediagrams-application-property-access.md)|
|[Count](alldatabasediagrams-count-property-access.md)|
|[Item](alldatabasediagrams-item-property-access.md)|
|[Parent](alldatabasediagrams-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
