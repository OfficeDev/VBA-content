---
title: ToggleButton.Value Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 7e935582-fcae-a703-4fbe-eda43852d0ce
ms.date: 06/08/2017
---


# ToggleButton.Value Property (Outlook Forms Script)

Returns or sets a  **Variant** that specifies whether the toggle button is selected. Read/write.


## Syntax

 _expression_. **Value**

 _expression_A variable that represents a  **ToggleButton** object.


## Remarks

The settings for  **Value** are:



|**Control**|**Description**|
|:-----|:-----|
|Null|Indicates the item is in a null state, neither selected nor cleared.|
|True| Indicates the item is selected.|
|False|Indicates the item is cleared.|
|Zero (0)|Indicates the first page. The maximum value is one less than the number of pages.|

