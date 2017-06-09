---
title: ColorStops.Application Property (Excel)
keywords: vbaxl10.chm852073
f1_keywords:
- vbaxl10.chm852073
ms.prod: excel
api_name:
- Excel.ColorStops.Application
ms.assetid: 68c43e6a-7e68-777d-67a0-a895db4d351d
ms.date: 06/08/2017
---


# ColorStops.Application Property (Excel)

When used without an object qualifier, this property returns an  **Application** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ An expression that returns a **ColorStops** object.


### Return Value

Application


## Example

This example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If 

```


## See also


#### Concepts


[ColorStops Object](colorstops-object-excel.md)

