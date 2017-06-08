---
title: DataLabels.Propagate Method (Excel)
keywords: vbaxl10.chm584110
f1_keywords:
- vbaxl10.chm584110
ms.prod: excel
ms.assetid: cf81fe7c-fb9c-bcd5-bd29-aef898c9c265
ms.date: 06/08/2017
---


# DataLabels.Propagate Method (Excel)

Enables you to take the contents and formatting of a single data label and apply it to every other data label on the series.


## Syntax

 _expression_ . **Propagate**_(Index)_

 _expression_ A variable that represents a **DataLabels** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|VARIANT|The index number of the data label to propagate.|

### Remarks

If the source data label supports fields that are incompatible with the destination data label, those fields will be inserted as their [FIELD NAME]. For example, if a data label on a pie series with a percentage field is propagated to a data label on a column chart, that percentage field will become an unresolved field showing [PERCENTAGE].


 **Note**  Passing an argument of zero (0) resets the data labels to your current prototype.


### Return value

 **VOID**


## See also


#### Concepts


[DataLabels Object](datalabels-object-excel.md)

