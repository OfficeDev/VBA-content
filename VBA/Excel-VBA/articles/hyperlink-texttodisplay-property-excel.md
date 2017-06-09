---
title: Hyperlink.TextToDisplay Property (Excel)
keywords: vbaxl10.chm536085
f1_keywords:
- vbaxl10.chm536085
ms.prod: excel
api_name:
- Excel.Hyperlink.TextToDisplay
ms.assetid: b7b8e4ef-2a37-1733-f9a0-2bd6e7367f8d
ms.date: 06/08/2017
---


# Hyperlink.TextToDisplay Property (Excel)

Returns or sets the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink. Read/write  **String** .


## Syntax

 _expression_ . **TextToDisplay**

 _expression_ A variable that represents a **Hyperlink** object.


## Example

This example sets the text to be displayed for the first hyperlink on the active worksheet.


```vb
ActiveSheet.Hyperlinks(1).TextToDisplay = _ 
 "Company Home Page"
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-excel.md)

