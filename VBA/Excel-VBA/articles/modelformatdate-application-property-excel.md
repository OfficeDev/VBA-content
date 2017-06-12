---
title: ModelFormatDate.Application Property (Excel)
keywords: vbaxl10.chm983073
f1_keywords:
- vbaxl10.chm983073
ms.assetid: a0932f28-60a5-34ce-a678-78cb72b33b31
ms.date: 06/08/2017
ms.prod: excel
---


# ModelFormatDate.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **ModelFormatDate** object.


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


#### Other resources


[ModelFormatDate Object](modelformatdate-object-excel.md)


