---
title: OlkListBox.SetSelected Method (Outlook)
keywords: vbaol11.chm1000271
f1_keywords:
- vbaol11.chm1000271
ms.prod: outlook
api_name:
- Outlook.OlkListBox.SetSelected
ms.assetid: ee8a6553-4cf4-b99d-9289-bec4d86e7c32
ms.date: 06/08/2017
---


# OlkListBox.SetSelected Method (Outlook)

Sets the selected state of an item at the specified location in the list to the given  _Selected_ value.


## Syntax

 _expression_ . **SetSelected**( **_Index_** , **_Selected_** )

 _expression_ A variable that represents an **OlkListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A zero-based value that specifies the location of an item in the list.|
| _Selected_|Required| **Boolean**| **True** to indicate that the item should be selected, **False** to indicate that the item should not be selected.|

## Remarks

If  _Index_ is outside the range of the allowed values (between zero and **[ListCount](olklistbox-listcount-property-outlook.md)** -1), then an out-of-bounds error will be returned.


## See also


#### Concepts


[OlkListBox Object](olklistbox-object-outlook.md)

