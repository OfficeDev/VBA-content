---
title: AllQueries Object (Access)
keywords: vbaac10.chm12689
f1_keywords:
- vbaac10.chm12689
ms.prod: access
api_name:
- Access.AllQueries
ms.assetid: 9b67f04c-2642-0dcc-2a64-8ca8fa7249b3
ms.date: 06/08/2017
---


# AllQueries Object (Access)

The  **AllQueries** collection contains an **[AccessObject](accessobject-object-access.md)** for each query in the **[CurrentData](currentdata-object-access.md)** or **[CodeData](codedata-object-access.md)** object.


## Remarks

The  **CurrentData** or **CodeData** object has an **AllQueries** collection containing **AccessObject** objects that describe instances of all queries specified by **CurrentData** or **CodeData**. For example, you can enumerate the **AllQueries** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual  **AccessObject** object in the **AllQueries** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllQueries** collection, it's better to refer to the query by name because a query's collection index may change.

The  **AllQueries** collection is indexed beginning with zero. If you refer to a query by its index, the first query is AllQueries(0), the second query is AllQueries(1), and so on.


 **Note**  


## Example

The following example prints the name of each open  **AccessObject** object in the **AllQueries** collection.


```
Sub AllQueries() 
    Dim obj As AccessObject, dbs As Object 
    Set dbs = Application.CurrentData 
    ' Search for open AccessObject objects in AllQueries collection. 
    For Each obj In dbs.AllQueries 
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
|[Application](allqueries-application-property-access.md)|
|[Count](allqueries-count-property-access.md)|
|[Item](allqueries-item-property-access.md)|
|[Parent](allqueries-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
