---
title: DefaultWebOptions.AllowPNG Property (Word)
keywords: vbawd10.chm165871626
f1_keywords:
- vbawd10.chm165871626
ms.prod: word
api_name:
- Word.DefaultWebOptions.AllowPNG
ms.assetid: 37cb4096-cd21-be5f-1f55-8786b56fc7a6
ms.date: 06/08/2017
---


# DefaultWebOptions.AllowPNG Property (Word)

 **False** if PNG (Portable Network Graphics) is not allowed as an output format. Read/write **Boolean** .


## Syntax

 _expression_ . **AllowPNG**

 _expression_ A variable that represents a **[DefaultWebOptions](defaultweboptions-object-word.md)** collection.


## Remarks

 **True** if PNG is allowed as an image format when you save a document as a Web page. The default value is **False** .

If you save images in the PNG format and if the Web browsers you are targeting support the PNG format, you might improve the image quality or reduce the size of those image files, and therefore decrease the download time.


## Example

This example enables PNG as an output format for the active document.


```vb
ActiveDocument.WebOptions.AllowPNG = True
```

Alternatively, PNG can be enabled as the global default for the application for newly created documents.




```vb
Application.DefaultWebOptions.AllowPNG = True
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

