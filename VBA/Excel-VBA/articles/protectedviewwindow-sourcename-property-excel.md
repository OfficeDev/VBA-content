---
title: ProtectedViewWindow.SourceName Property (Excel)
keywords: vbaxl10.chm914081
f1_keywords:
- vbaxl10.chm914081
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.SourceName
ms.assetid: e5347e6e-b9d4-d3b1-ca41-ba577d836e31
ms.date: 06/08/2017
---


# ProtectedViewWindow.SourceName Property (Excel)

Returns the name of the source file that is open in the specified  **Protected View** window. Read-only


## Syntax

 _expression_ . **SourceName**

 _expression_ A variable that represents a **[ProtectedViewWindow](protectedviewwindow-object-excel.md)** object.


### Return Value

 **String**


## Remarks

This property does not return the path for the source file. To return the path, use the  **[SourcePath](protectedviewwindow-sourcepath-property-excel.md)** property of the **ProtectedViewWindow** object.


## Example

The following example returns the path and name of the workbook associated with the specified  **Protected View** window.


```vb
MsgBox ActiveProtectedViewWindow.SourcePath &; "\" _ 
 &; ActiveProtectedViewWindow.SourceName
```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-excel.md)

