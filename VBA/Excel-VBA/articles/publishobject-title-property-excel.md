---
title: PublishObject.Title Property (Excel)
keywords: vbaxl10.chm652080
f1_keywords:
- vbaxl10.chm652080
ms.prod: excel
api_name:
- Excel.PublishObject.Title
ms.assetid: 3e8eae5c-62f5-3d72-2c27-ff5107153adc
ms.date: 06/08/2017
---


# PublishObject.Title Property (Excel)

Returns or sets the title of the Web page when the document is saved as a Web page. Read/write  **String** .


## Syntax

 _expression_ . **Title**

 _expression_ A variable that represents a **PublishObject** object.


## Remarks

The title is usually displayed in the window title bar when the document is viewed in the Web browser.


## Example

This example sets the Web page title to "Sales Forecast" when the first item in the first workbook is saved as a Web page.


```vb
Workbooks(1).PublishObjects(1).Title = "Sales Forecast"
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-excel.md)

