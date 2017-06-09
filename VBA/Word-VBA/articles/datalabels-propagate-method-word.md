---
title: DataLabels.Propagate Method (Word)
keywords: vbawd10.chm207489004
f1_keywords:
- vbawd10.chm207489004
ms.prod: word
ms.assetid: 72885eed-605b-70f1-386d-9fdf2c40ef9d
ms.date: 06/08/2017
---


# DataLabels.Propagate Method (Word)

Propagates the contents and formatting of the specified data label to all the other data labels in the series.


## Syntax

 _expression_ . **Propagate**_(Index)_

 _expression_ A variable that represents a **DataLabels** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Index_|Required|VARIANT|The index number in the  **DataLabels** collection of the data label to propagate.|

### Return value

 **VOID**


## Remarks

If the source data label supports fields that are incompatible with the destination data label, those fields are inserted in the form [ _Field Name_ ].


## See also


#### Concepts


[DataLabels Collection](datalabels-object-word.md)

