---
title: Adjustments.Application Property (Excel)
ms.prod: excel
api_name:
- Excel.Adjustments.Application
ms.assetid: 2875f3fa-d584-2ba5-c445-ac4dbad25af2
ms.date: 06/08/2017
---


# Adjustments.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents an **Adjustments** object.


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


[Adjustments Object](adjustments-object-excel.md)

