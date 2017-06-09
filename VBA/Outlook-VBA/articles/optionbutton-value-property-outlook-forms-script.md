---
title: OptionButton.Value Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 2ab8a0e5-2b82-5542-3343-2b4599141ef8
ms.date: 06/08/2017
---


# OptionButton.Value Property (Outlook Forms Script)

Returns or sets a  **Variant** that specifies whether the option button is selected. Read/write.


## Syntax

 _expression_. **Value**

 _expression_A variable that represents an  **OptionButton** object.


## Remarks

The settings for  **Value** are:



|**Control**|**Description**|
|:-----|:-----|
|Null|Indicates the item is in a null state, neither selected nor cleared.|
|True| Indicates the item is selected.|
|False|Indicates the item is cleared.|
|Zero (0)|Indicates the first page. The maximum value is one less than the number of pages.|

