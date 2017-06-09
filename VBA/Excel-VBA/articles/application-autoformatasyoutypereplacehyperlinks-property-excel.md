---
title: Application.AutoFormatAsYouTypeReplaceHyperlinks Property (Excel)
keywords: vbaxl10.chm133281
f1_keywords:
- vbaxl10.chm133281
ms.prod: excel
api_name:
- Excel.Application.AutoFormatAsYouTypeReplaceHyperlinks
ms.assetid: 92c02556-f39a-7ca4-31f5-88a5c9da12ea
ms.date: 06/08/2017
---


# Application.AutoFormatAsYouTypeReplaceHyperlinks Property (Excel)

 **True** (default) if Microsoft Excel automatically formats hyperlinks as you type. **False** if Excel does not automatically format hyperlinks as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplaceHyperlinks**

 _expression_ A variable that represents an **Application** object.


## Example

In this example, Microsoft Excel determines if the ability to format hyperlinks automatically as they are typed is enabled and notifies the user.


```vb
Sub CheckHyperlinks() 
 
 ' Determine if automatic formatting is enabled and notify user. 
 If Application.AutoFormatAsYouTypeReplaceHyperlinks = True Then 
 MsgBox "Automatic formatting for typing in hyperlinks is enabled." 
 Else 
 MsgBox "Automatic formatting for typing in hyperlinks is not enabled." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

