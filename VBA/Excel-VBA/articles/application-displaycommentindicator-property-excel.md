---
title: Application.DisplayCommentIndicator Property (Excel)
keywords: vbaxl10.chm133123
f1_keywords:
- vbaxl10.chm133123
ms.prod: excel
api_name:
- Excel.Application.DisplayCommentIndicator
ms.assetid: 8617da4e-97cb-fe57-bb51-a9c671e2ff27
ms.date: 06/08/2017
---


# Application.DisplayCommentIndicator Property (Excel)

Returns or sets the way cells display comments and indicators. Can be one of the  **[XlCommentDisplayMode](xlcommentdisplaymode-enumeration-excel.md)** constants.


## Syntax

 _expression_ . **DisplayCommentIndicator**

 _expression_ A variable that represents an **Application** object.


## Example

This example hides cell tips but retains comment indicators.


```vb
Application.DisplayCommentIndicator = xlCommentIndicatorOnly
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

