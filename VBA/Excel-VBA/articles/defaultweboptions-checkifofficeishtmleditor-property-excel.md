---
title: DefaultWebOptions.CheckIfOfficeIsHTMLEditor Property (Excel)
keywords: vbaxl10.chm660079
f1_keywords:
- vbaxl10.chm660079
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.CheckIfOfficeIsHTMLEditor
ms.assetid: 29b77ad1-11ea-f930-a4ab-6bb957287eea
ms.date: 06/08/2017
---


# DefaultWebOptions.CheckIfOfficeIsHTMLEditor Property (Excel)

 **True** if Microsoft Excel checks to see whether an Office application is the default HTML editor when you start Excel. **False** if Excel does not perform this check. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **CheckIfOfficeIsHTMLEditor**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Remarks

This property is used only if the Web browser you are using supports HTML editing and HTML editors.

To use a different HTML editor, you must set this property to  **False** and then register the editor as the default system HTML editor.


## Example

This example causes Microsoft Excel not to check to see whether it is the default HTML editor.


```vb
Application.DefaultWebOptions.CheckIfOfficeIsHTMLEditor = False
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

