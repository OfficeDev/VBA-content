---
title: WebOptions.AllowPNG Property (Excel)
keywords: vbaxl10.chm662078
f1_keywords:
- vbaxl10.chm662078
ms.prod: excel
api_name:
- Excel.WebOptions.AllowPNG
ms.assetid: 4fad6401-af54-ad7f-a46f-8110e8c00ad4
ms.date: 06/08/2017
---


# WebOptions.AllowPNG Property (Excel)

 **True** if PNG (Portable Network Graphics) is allowed as an image format when you save documents as a Web page. **False** if PNG is not allowed as an output format. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **AllowPNG**

 _expression_ A variable that represents a **WebOptions** object.


## Remarks

If you save images in the PNG format as opposed to any other file format, you might improve the image quality or reduce the size of those image files, and therefore decrease the download time, assuming that the Web browsers you are targeting support the PNG format.


## Example

This example enables PNG as an output format for the first workbook.


```vb
Application.Workbooks(1).WebOptions.AllowPNG = True
```

Alternatively, PNG can be enabled as the global default for the application for newly created documents.




```vb
Application.DefaultWebOptions.AllowPNG = True
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-excel.md)

