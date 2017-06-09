---
title: Hyperlink.SubAddress Property (Excel)
keywords: vbaxl10.chm536077
f1_keywords:
- vbaxl10.chm536077
ms.prod: excel
api_name:
- Excel.Hyperlink.SubAddress
ms.assetid: e83633c1-66b7-02f1-0e05-0397dc4f41ae
ms.date: 06/08/2017
---


# Hyperlink.SubAddress Property (Excel)

Returns or sets the location within the document associated with the hyperlink. Read/write  **String** .


## Syntax

 _expression_ . **SubAddress**

 _expression_ A variable that represents a **Hyperlink** object.


## Example

This example topic adds a range location to the hyperlink for shape one.


```vb
Worksheets(1).Shapes(1).Hyperlink.SubAddress = "A1:B10"
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-excel.md)

