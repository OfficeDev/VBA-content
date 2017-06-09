---
title: Application.WrapText Method (Project)
keywords: vbapj.chm708
f1_keywords:
- vbapj.chm708
ms.prod: project-server
api_name:
- Project.Application.WrapText
ms.assetid: 0aaabac2-ee1d-694c-45ac-f522a0034724
ms.date: 06/08/2017
---


# Application.WrapText Method (Project)

Toggles the  **Wrap Text** setting in a column.


## Syntax

 _expression_. **WrapText**( ** _Column_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Column_|Optional|**Integer**|The target column identifier. If omitted, the  **WrapText** method is applied to the column containing the active cell.|

### Return Value

 **Boolean**


## Remarks

The  **WrapText** method corresponds to the **Wrap Text** command in the option menu for a column.


