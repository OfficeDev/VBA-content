---
title: DefaultWebOptions.Encoding Property (Excel)
keywords: vbaxl10.chm660086
f1_keywords:
- vbaxl10.chm660086
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.Encoding
ms.assetid: 53164ab3-b0f5-ed8e-76f8-840cbd8e23bc
ms.date: 06/08/2017
---


# DefaultWebOptions.Encoding Property (Excel)

Returns or sets the document encoding (code page or character set) to be used by the Web browser when you view the saved document. The default is the system code page. Read/write  **[MsoEncoding](http://msdn.microsoft.com/library/286bed6e-6028-a252-5e4f-b505234d9d34%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **Encoding**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Remarks

You cannot use any of the constants that have the suffix  **AutoDetect** . These constants are used by the **[ReloadAs](workbook-reloadas-method-excel.md)** method.


## Example

This example checks to see whether the default document encoding is Western, and then it sets the string  `strDocEncoding` accordingly.


```vb
If Application.DefaultWebOptions.Encoding = msoEncodingWestern Then 
    strDocEncoding = "Western" 
Else 
    strDocEncoding = "Other" 
End If
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

