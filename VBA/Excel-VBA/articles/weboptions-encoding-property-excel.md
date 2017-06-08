---
title: WebOptions.Encoding Property (Excel)
keywords: vbaxl10.chm662082
f1_keywords:
- vbaxl10.chm662082
ms.prod: excel
api_name:
- Excel.WebOptions.Encoding
ms.assetid: 99395ad8-4503-eac2-b194-6a4706e5264d
ms.date: 06/08/2017
---


# WebOptions.Encoding Property (Excel)

Returns or sets the document encoding (code page or character set) to be used by the Web browser when you view the saved document. The default is the system code page. Read/write  **[MsoEncoding](http://msdn.microsoft.com/library/286bed6e-6028-a252-5e4f-b505234d9d34%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **Encoding**

 _expression_ A variable that represents a **WebOptions** object.


## Remarks

You cannot use any of the constants that have the suffix  **AutoDetect** . These constants are used by the **[ReloadAs](workbook-reloadas-method-excel.md)** method.


## See also


#### Concepts


[WebOptions Object](weboptions-object-excel.md)

