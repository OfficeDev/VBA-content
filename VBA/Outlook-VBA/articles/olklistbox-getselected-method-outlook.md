---
title: OlkListBox.GetSelected Method (Outlook)
keywords: vbaol11.chm1000270
f1_keywords:
- vbaol11.chm1000270
ms.prod: outlook
api_name:
- Outlook.OlkListBox.GetSelected
ms.assetid: f1af9a89-09aa-79da-ebbf-bce0948b4427
ms.date: 06/08/2017
---


# OlkListBox.GetSelected Method (Outlook)

Returns a  **Boolean** that indicates if the indexed item is currently selected.


## Syntax

 _expression_ . **GetSelected**( **_Index_** )

 _expression_ A variable that represents an **OlkListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A zero-based value that specifies the location of an item in the list.|

### Return Value

A  **Boolean** value that is **True** if the specified item is currently selected, **False** otherwise.


## Remarks

If  _Index_ is outside the range of the allowed values (between zero and **[ListCount](olklistbox-listcount-property-outlook.md)** -1), then an out-of-bounds error will be returned.


## See also


#### Concepts


[OlkListBox Object](olklistbox-object-outlook.md)

