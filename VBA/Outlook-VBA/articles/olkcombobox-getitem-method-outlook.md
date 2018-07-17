---
title: OlkComboBox.GetItem Method (Outlook)
keywords: vbaol11.chm1000224
f1_keywords:
- vbaol11.chm1000224
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.GetItem
ms.assetid: 650fa823-fbb9-9013-86af-4f55367475c3
ms.date: 06/08/2017
---


# OlkComboBox.GetItem Method (Outlook)

Obtains a  **String** that represents an item at the specified location in the list of the combo box control.


## Syntax

 _expression_ . **GetItem**( **_Index_** )

 _expression_ A variable that represents an **OlkComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A zero-based value that specifies the location of an item in the list.|

### Return Value

A  **String** value that represents the item at the specified location in the list.


## Remarks

If  _Index_ is outside the range of the allowed values (between zero and **[ListCount](olkcombobox-listcount-property-outlook.md)** -1), then an out-of-bounds error will be returned.


## See also


#### Concepts


[OlkComboBox Object](olkcombobox-object-outlook.md)

