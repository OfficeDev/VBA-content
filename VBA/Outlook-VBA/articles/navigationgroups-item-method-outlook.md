---
title: NavigationGroups.Item Method (Outlook)
keywords: vbaol11.chm2857
f1_keywords:
- vbaol11.chm2857
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.Item
ms.assetid: a6521179-fa65-b5af-629a-458a852a29b4
ms.date: 06/08/2017
---


# NavigationGroups.Item Method (Outlook)

Returns a  **[NavigationGroup](navigationgroup-object-outlook.md)** object from the collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **NavigationGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number of the object.|

### Return Value

A  **NavigationGroup** object that represents the specified object.


## Remarks

The index value of a  **NavigationGroup** in the collection represents the ordinal position of the navigation group when displayed in the Navigation Pane. Changing the position of navigation groups also changes the index values of navigation groups contained within the **[NavigationGroups](navigationgroups-object-outlook.md)** collection.


## See also


#### Concepts


[NavigationGroups Object](navigationgroups-object-outlook.md)

