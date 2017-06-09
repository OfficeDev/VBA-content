---
title: LineFormat.Application Property (Excel)
ms.prod: excel
api_name:
- Excel.LineFormat.Application
ms.assetid: c90f22c9-b9e5-a91c-23fb-3301b709000a
ms.date: 06/08/2017
---


# LineFormat.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **LineFormat** object.


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


[LineFormat Object](lineformat-object-excel.md)

