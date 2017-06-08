---
title: AllModules Object (Access)
keywords: vbaac10.chm12686
f1_keywords:
- vbaac10.chm12686
ms.prod: access
api_name:
- Access.AllModules
ms.assetid: 322815ae-3afd-f299-0ce9-2e9dbbb8536a
ms.date: 06/08/2017
---


# AllModules Object (Access)

The  **AllModules** collection contains an **[AccessObject](accessobject-object-access.md)** of each module in the **[CurrentProject](currentproject-object-access.md)** or **[CodeProject](codeproject-object-access.md)** object.


## Remarks

The  **CurrentProject** or **CodeProject** object has an **AllModules** collection containing **AccessObject** objects that describe instances of all the **Module** objects specified by **CurrentProject** or **CodeProject**. For example, you can enumerate the **AllModules** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual  **AccessObject** object in the **AllModules** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllModules** collection, it's better to refer to the module by name because a module's collection index may change.

The  **AllModules** collection is indexed beginning with zero. If you refer to a module by its index, the first module is AllModules(0), the second module is AllModules(1), and so on.


 **Note**   To list all open modules in the database, use the **[IsLoaded](accessobject-isloaded-property-access.md)** property of each **AccessObject** object in the **AllModules** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a module.

You can't add or delete an  **AccessObject** object from the **AllModules** collection.


## Example

The following example prints the name of each open  **AccessObject** object in the **AllModules** collection.


```
Sub AllModules() 
 Dim obj As AccessObject, dbs As Object 
 Set dbs = Application.CurrentProject 
 ' Search for open AccessObject objects in AllModules collection. 
 For Each obj In dbs.AllModules 
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
|[Application](allmodules-application-property-access.md)|
|[Count](allmodules-count-property-access.md)|
|[Item](allmodules-item-property-access.md)|
|[Parent](allmodules-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
