---
title: PublishObject.Filename Property (Excel)
keywords: vbaxl10.chm652081
f1_keywords:
- vbaxl10.chm652081
ms.prod: excel
api_name:
- Excel.PublishObject.Filename
ms.assetid: bd0a4a76-62b8-95bc-37d3-efc1249f9bc8
ms.date: 06/08/2017
---


# PublishObject.Filename Property (Excel)

Returns or sets the URL (on the intranet or the Web) or path (local or network) to the location where the specified source object was saved. Read/write  **String** .


## Syntax

 _expression_ . **Filename**

 _expression_ A variable that represents a **PublishObject** object.


## Remarks

The  **FileName** property generates an error if a folder in the specified path doesn't exist.


## Example

This example sets the location where the first item in the active workbook is to be saved.


```vb
ActiveWorkbook.PublishObjects(1).FileName = _ 
 "\\Server2\Q1\StockReport.htm"
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-excel.md)

