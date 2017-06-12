---
title: XmlNamespaces.Application Property (Excel)
keywords: vbaxl10.chm745073
f1_keywords:
- vbaxl10.chm745073
ms.prod: excel
api_name:
- Excel.XmlNamespaces.Application
ms.assetid: c885b4c1-429d-8318-5898-8f1a881163ba
ms.date: 06/08/2017
---


# XmlNamespaces.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **XmlNamespaces** object.


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


[XmlNamespaces Object](xmlnamespaces-object-excel.md)

