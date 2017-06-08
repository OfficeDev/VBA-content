---
title: DefaultWebOptions.AllowPNG Property (Excel)
keywords: vbaxl10.chm660082
f1_keywords:
- vbaxl10.chm660082
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.AllowPNG
ms.assetid: b4cdab42-25ed-e313-ebf2-fc9abd68474b
ms.date: 06/08/2017
---


# DefaultWebOptions.AllowPNG Property (Excel)

 **True** if PNG (Portable Network Graphics) is allowed as an image format when you save documents as a Web page. **False** if PNG is not allowed as an output format. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **AllowPNG**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Remarks

If you save images in the PNG format as opposed to any other file format, you might improve the image quality or reduce the size of those image files, and therefore decrease the download time, assuming that the Web browsers you are targeting support the PNG format.


## Example

Alternatively, PNG can be enabled as the global default for the application for newly created documents.


```vb
Application.DefaultWebOptions.AllowPNG = True
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

