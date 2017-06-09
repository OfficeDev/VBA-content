---
title: AllViews Object (Access)
keywords: vbaac10.chm12690
f1_keywords:
- vbaac10.chm12690
ms.prod: access
api_name:
- Access.AllViews
ms.assetid: f56bee24-a972-fbdf-f74a-0ac83825e3bb
ms.date: 06/08/2017
---


# AllViews Object (Access)

The  **AllViews** collection contains an **[AccessObject](accessobject-object-access.md)** for each view in the **[CurrentData](currentdata-object-access.md)** or **[CodeData](codedata-object-access.md)** object.


## Remarks

The  **CurrentData** or **CodeData** object has an **AllViews** collection containing **AccessObject** objects that describe instances of all views specified by **CurrentData** or **CodeData**. For example, you can enumerate the **AllViews** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual  **AccessObject** object in the **AllViews** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllViews** collection, it's better to refer to the view by name because a view's collection index may change.

The  **AllViews** collection is indexed beginning with zero. If you refer to a view by its index, the first view is AllViews(0), the second table is AllViews(1), and so on.


 **Note**  


## Example

The following example prints the name of each open  **AccessObject** object in the **AllViews** collection.


```
Sub AllViews() 
    Dim obj As AccessObject, dbs As Object 
    Set dbs = Application.CurrentData 
    ' Search for open AccessObject objects in AllViews collection. 
    For Each obj In dbs.AllViews 
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
|[Application](allviews-application-property-access.md)|
|[Count](allviews-count-property-access.md)|
|[Item](allviews-item-property-access.md)|
|[Parent](allviews-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
