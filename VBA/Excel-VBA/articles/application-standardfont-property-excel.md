---
title: Application.StandardFont Property (Excel)
keywords: vbaxl10.chm133210
f1_keywords:
- vbaxl10.chm133210
ms.prod: excel
api_name:
- Excel.Application.StandardFont
ms.assetid: 6bde5ec0-8868-fa00-52e3-b7387f39f56d
ms.date: 06/08/2017
---


# Application.StandardFont Property (Excel)

Returns or sets the name of the standard font. Read/write  **String** .


## Syntax

 _expression_ . **StandardFont**

 _expression_ A variable that represents an **Application** object.


## Remarks

If you change the standard font by using this property, the change doesn't take effect until you restart Microsoft Excel.


## Example

This example sets the standard font to Geneva (on the Macintosh) or Arial (in Windows).


```vb
If Application.OperatingSystem Like "*Macintosh*" Then 
 Application.StandardFont = "Geneva" 
Else 
 Application.StandardFont = "Arial" 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

