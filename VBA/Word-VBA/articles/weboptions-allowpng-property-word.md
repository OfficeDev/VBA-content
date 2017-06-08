---
title: WebOptions.AllowPNG Property (Word)
keywords: vbawd10.chm165937159
f1_keywords:
- vbawd10.chm165937159
ms.prod: word
api_name:
- Word.WebOptions.AllowPNG
ms.assetid: 61fb3c31-0c6a-f4f0-390b-81d0ffa348ec
ms.date: 06/08/2017
---


# WebOptions.AllowPNG Property (Word)

 **True** if PNG (Portable Network Graphics) is allowed as an image format when you save a document as a Web page. **False** if PNG is not allowed as an output format. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **AllowPNG**

 _expression_ A variable that represents a **[WebOptions](weboptions-object-word.md)** collection.


## Remarks

If you save images in the PNG format and if the Web browsers you are targeting support the PNG format, you might improve the image quality or reduce the size of those image files, and therefore decrease the download time.


## Example

This example enables PNG as an output format for the active document.


```vb
ActiveDocument.WebOptions.AllowPNG = True
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-word.md)

