---
title: Application.IsSandboxed Property (Excel)
keywords: vbaxl10.chm133332
f1_keywords:
- vbaxl10.chm133332
ms.prod: excel
api_name:
- Excel.Application.IsSandboxed
ms.assetid: d5a40aa3-470b-7861-691f-de418d13bd8b
ms.date: 06/08/2017
---


# Application.IsSandboxed Property (Excel)

Returns  **True** if the specified workbook is open in a **Protected View** window. Read-only


## Syntax

 _expression_ . **IsSandboxed**

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

Use the  **IsSandboxed** property to determine if a workbook is open in a **Protected View** window.


## Example

The following code example displays whether the specified workbook is open in a  **Protected View** window.


```vb
Sub CheckIfSandboxed(wbk As Workbook) 
 MsgBox wbk.Application.IsSandboxed 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

