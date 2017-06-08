---
title: ColorFormat.Application Property (Excel)
ms.prod: excel
api_name:
- Excel.ColorFormat.Application
ms.assetid: e9b68987-dceb-8bd6-13af-be60076e3e73
ms.date: 06/08/2017
---


# ColorFormat.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **ColorFormat** object.


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


[ColorFormat Object](colorformat-object-excel.md)

