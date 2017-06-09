---
title: Modules Object (Access)
keywords: vbaac10.chm12289
f1_keywords:
- vbaac10.chm12289
ms.prod: access
api_name:
- Access.Modules
ms.assetid: f60a9929-4b79-cfed-8fb3-a4869a3afe9f
ms.date: 06/08/2017
---


# Modules Object (Access)

The  **Modules** collection contains all open standard modules and class modules in a Microsoft Access database.


## Remarks

All open modules are included in the  **Modules** collection, whether they are uncompiled, are compiled, are in break mode, or contain the code that's running.

 To determine whether an individual **Module** object represents a standard module or a class module, check the **Module** object's **Type** property.

The  **Modules** collection belongs to the Microsoft Access **Application** object.

Individual  **Module** objects in the **Modules** collection are indexed beginning with zero.


## Example

The following example illustrates how to use the Modules collection to loop through the open modules. The example prints the name if each open module in the Immediate window.


```
 
Sub PrintOpenModuleNames() 
 Dim i As Integer 
 Dim modOpenModules As Modules 
 
 Set modOpenModules = Application.Modules 
 
 For i = 0 To modOpenModules.Count - 1 
 
 Debug.Print modOpenModules(i).Name 
 
 Next 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](modules-application-property-access.md)|
|[Count](modules-count-property-access.md)|
|[Item](modules-item-property-access.md)|
|[Parent](modules-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
