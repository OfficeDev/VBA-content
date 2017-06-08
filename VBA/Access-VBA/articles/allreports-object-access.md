---
title: AllReports Object (Access)
keywords: vbaac10.chm12684
f1_keywords:
- vbaac10.chm12684
ms.prod: access
api_name:
- Access.AllReports
ms.assetid: 5846cf60-41b4-e9f8-ea27-b9400a6d3861
ms.date: 06/08/2017
---


# AllReports Object (Access)

The  **AllReports** collection contains an **[AccessObject](accessobject-object-access.md)** for each report in the **[CurrentProject](currentproject-object-access.md)** or **[CodeProject](codeproject-object-access.md)** object.


## Remarks

The  **CurrentProject** or **CodeProject** object has an **AllReports** collection containing **AccessObject** objects that describe instances of all the reports in the database. For example, you can enumerate the **AllReports** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual  **AccessObject** object in the **AllReports** collection either by referring to the item by name, or by referring to its index within the collection. If you want to refer to a specific report in the **AllReports** collection, it's better to refer to the item by name because the index may change.

The  **AllReports** collection is indexed beginning with zero. If you refer to a report by its index, the first report is AllReports(0), the second report is AllReports(1), and so on.


 **Note**  To list all open reports in the database, use the  **[IsLoaded](accessobject-isloaded-property-access.md)** property of each **AccessObject** object in the **AllReports** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a report.

You can't add or delete an  **AccessObject** object from the **AllReports** collection.


## Example

The following example prints the name of each open  **AccessObject** object in the **AllReports** collection.


```
Sub AllReports() 
 Dim obj As AccessObject, dbs As Object 
 Set dbs = Application.CurrentProject 
 ' Search for open AccessObject objects in AllReports collection. 
 For Each obj In dbs.AllReports 
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
|[Application](allreports-application-property-access.md)|
|[Count](allreports-count-property-access.md)|
|[Item](allreports-item-property-access.md)|
|[Parent](allreports-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
