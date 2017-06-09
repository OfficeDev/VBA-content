---
title: OlkComboBox.SetItem Method (Outlook)
keywords: vbaol11.chm1000225
f1_keywords:
- vbaol11.chm1000225
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.SetItem
ms.assetid: 00cc1630-1423-5244-557b-acb2861401bf
ms.date: 06/08/2017
---


# OlkComboBox.SetItem Method (Outlook)

Sets the item at the specified location in the list of the combo box to the specified value.


## Syntax

 _expression_ . **SetItem**( **_Index_** , **_Item_** )

 _expression_ A variable that represents an **OlkComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A zero-based value that specifies the location of an item in the list.|
| _Item_|Required| **String**|The value to be used to update the list at the specified location.|

## Remarks

If  _Index_ is outside the range of the allowed values (between zero and **[ListCount](olklistbox-listcount-property-outlook.md)** -1), then an out-of-bounds error will be returned.


## See also


#### Concepts


[OlkComboBox Object](olkcombobox-object-outlook.md)

