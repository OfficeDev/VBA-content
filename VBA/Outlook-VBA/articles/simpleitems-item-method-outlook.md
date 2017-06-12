---
title: SimpleItems.Item Method (Outlook)
keywords: vbaol11.chm3398
f1_keywords:
- vbaol11.chm3398
ms.prod: outlook
api_name:
- Outlook.SimpleItems.Item
ms.assetid: 0b56d8a7-2bf5-a2e2-a269-b2d7377d2901
ms.date: 06/08/2017
---


# SimpleItems.Item Method (Outlook)

Returns an item in the  **[SimpleItems](simpleitems-object-outlook.md)** collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **SimpleItems** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The zero-based index number of the object in the  **SimpleItems** collection.|

### Return Value

An  **Object** that represents an Outlook item in the **SimpleItems** collection.


## Remarks

If this method fails to return an object in the collection as specified by the  _Index_ parameter, the method returns **Null** ( **Nothing** in Visual Basic).


## See also


#### Concepts


[SimpleItems Object](simpleitems-object-outlook.md)

