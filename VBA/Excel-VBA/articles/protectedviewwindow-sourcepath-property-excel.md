---
title: ProtectedViewWindow.SourcePath Property (Excel)
keywords: vbaxl10.chm914082
f1_keywords:
- vbaxl10.chm914082
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.SourcePath
ms.assetid: add00cce-b8e9-5a11-b1cb-27ac63798491
ms.date: 06/08/2017
---


# ProtectedViewWindow.SourcePath Property (Excel)

Returns the path of the source file that is open in the specified  **Protected View** window. Read-only


## Syntax

 _expression_ . **SourcePath**

 _expression_ A variable that represents a **[ProtectedViewWindow](protectedviewwindow-object-excel.md)** object.


### Return Value

 **String**


## Remarks

The path does not include a trailing character (for example, "C:\MSOffice"). Use the  **[PathSeparator](application-pathseparator-property-excel.md)** property to add the character that separates folders and drive letters. Use the **[SourceName](protectedviewwindow-sourcename-property-excel.md)** of the **ProtectedViewWindow** object to return the source file name without the path.


## Example


```vb
MsgBox ActiveProtectedViewWindow.SourcePath &; Application.PathSeparator _ 
 &; ActiveProtectedViewWindow.SourceName 

```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-excel.md)

