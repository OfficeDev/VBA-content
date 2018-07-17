---
title: ProtectedViewWindow.Workbook Property (Excel)
keywords: vbaxl10.chm914084
f1_keywords:
- vbaxl10.chm914084
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.Workbook
ms.assetid: 379b98f0-b177-7910-4968-ce4ed2f1ca9d
ms.date: 06/08/2017
---


# ProtectedViewWindow.Workbook Property (Excel)

Returns an object that represents the workbook that is open in the specified  **Protected View** window. Read-only


## Syntax

 _expression_ . **Workbook**

 _expression_ A variable that represents a **[ProtectedViewWindow](protectedviewwindow-object-excel.md)** object.


### Return Value

 **[Workbook](workbook-object-excel.md)**


## Remarks

Because a  **Protected View** window is designed to protect the user from potentially malicious code, the operations you can perform by using a **Workbook** object returned by the **Workbook** method will be limited. Any operation that is not allowed will return an error.

A workbook displayed in a protected view window is not a member of the  **[Workbooks](workbooks-object-excel.md)** collection. Instead, use the **Workbook** property of the **ProtectedViewWindow** object to access a workbook that is displayed in a protected view window.


## Example

 The following example uses the **Workbook** property to return the workbook that is open in the first **Protected View** window.


```vb
Dim wbProtected As Workbook 
 
If Application.ProtectedViewWindows.Count > 0 Then 
    Set wbProtected = Application.ProtectedViewWindows(1).Workbook 
End If 

```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-excel.md)

