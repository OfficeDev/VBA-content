---
title: Chart.BeforeRightClick Event (Excel)
keywords: vbaxl10.chm500079
f1_keywords:
- vbaxl10.chm500079
ms.prod: excel
api_name:
- Excel.Chart.BeforeRightClick
ms.assetid: d01f6911-2f6b-3118-27a2-dfafa48791ab
ms.date: 06/08/2017
---


# Chart.BeforeRightClick Event (Excel)

Occurs when a chart element is right-clicked, before the default right-click action.


## Syntax

 _expression_ . **BeforeRightClick**( **_Cancel_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the default right-click action isn't performed when the procedure is finished.|

## Remarks

Like other worksheet events, this event doesn't occur if you right-click while the pointer is on a shape or a command bar (a toolbar or menu bar).


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

