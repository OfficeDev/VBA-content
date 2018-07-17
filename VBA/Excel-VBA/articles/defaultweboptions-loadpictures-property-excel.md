---
title: DefaultWebOptions.LoadPictures Property (Excel)
keywords: vbaxl10.chm660075
f1_keywords:
- vbaxl10.chm660075
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.LoadPictures
ms.assetid: dc2bcc24-f30b-6d63-e85c-20f29a2aaf03
ms.date: 06/08/2017
---


# DefaultWebOptions.LoadPictures Property (Excel)

 **True** if images are loaded when you open a document in Microsoft Excel, usually when the images and document were not created in Microsoft Excel. **False** if the images are not loaded. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **LoadPictures**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Example

This example causes images to load when the document is opened in Excel.


```vb
Application.DefaultWebOptions.LoadPictures = True
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

