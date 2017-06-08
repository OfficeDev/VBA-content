---
title: Application.MouseAvailable Property (Excel)
keywords: vbaxl10.chm133167
f1_keywords:
- vbaxl10.chm133167
ms.prod: excel
api_name:
- Excel.Application.MouseAvailable
ms.assetid: b22f9d44-6a84-6716-d663-450f08c5557d
ms.date: 06/08/2017
---


# Application.MouseAvailable Property (Excel)

 **True** if a mouse is available. Read-only **Boolean** .


## Syntax

 _expression_ . **MouseAvailable**

 _expression_ A variable that represents an **Application** object.


## Example

This example displays a message if a mouse isn't available.


```vb
If Application.MouseAvailable = False Then 
 MsgBox "Your system does not have a mouse" 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

