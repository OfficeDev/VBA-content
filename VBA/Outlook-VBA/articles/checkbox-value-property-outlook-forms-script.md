---
title: CheckBox.Value Property (Outlook Forms Script)
keywords: olfm10.chm2002180
f1_keywords:
- olfm10.chm2002180
ms.prod: outlook
ms.assetid: 24b3b4ab-e7cc-f024-c8b4-32db5dd389c7
ms.date: 06/08/2017
---


# CheckBox.Value Property (Outlook Forms Script)

Returns or sets a  **Variant** that specifies whether the check box is selected. Read/write.


## Syntax

 _expression_. **Value**

 _expression_A variable that represents a  **CheckBox** object.


## Remarks

The settings for  **Value** are:



|**Control**|**Description**|
|:-----|:-----|
| **Null**|Indicates the item is in a null state, neither selected nor cleared.|
| **True**| Indicates the item is selected.|
| **False**|Indicates the item is cleared.|
| **Zero** (0)|Indicates the first page. The maximum value is one less than the number of pages.|

