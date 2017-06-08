---
title: Application.Path Property (Excel)
keywords: vbaxl10.chm133189
f1_keywords:
- vbaxl10.chm133189
ms.prod: excel
api_name:
- Excel.Application.Path
ms.assetid: 0ef5d0fc-f46a-c133-232a-8a20cf2d4034
ms.date: 06/08/2017
---


# Application.Path Property (Excel)

Returns a  **String** value that represents the complete path to the application, excluding the final separator and name of the application.


## Syntax

 _expression_ . **Path**

 _expression_ A variable that represents an **Application** object.


## Example

This example displays the complete path to Microsoft Excel.


```vb
Sub TotalPath() 
 
 MsgBox "The path is " &; Application.Path 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

