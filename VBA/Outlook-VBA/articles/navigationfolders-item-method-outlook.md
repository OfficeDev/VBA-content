---
title: NavigationFolders.Item Method (Outlook)
keywords: vbaol11.chm2896
f1_keywords:
- vbaol11.chm2896
ms.prod: outlook
api_name:
- Outlook.NavigationFolders.Item
ms.assetid: 1688b2ef-a4a1-fc8a-513e-0d5e234f10dd
ms.date: 06/08/2017
---


# NavigationFolders.Item Method (Outlook)

Returns a  **[NavigationFolder](navigationfolder-object-outlook.md)** object from the collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **NavigationFolders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either the index number of the object, or the value used to match the default property of an object in the collection.|

### Return Value

A  **NavigationFolder** object that represents the specified object.


## Remarks

The index value of a  **NavigationFolder** in the collection represents the ordinal position of the folder in the navigation group when displayed in the Navigation Pane. Changing the position of navigation folders within the navigation group also changes the index values of folders contained within the **[NavigationFolders](navigationfolders-object-outlook.md)** collection for that **[NavigationGroup](navigationgroup-object-outlook.md)** object.


## See also


#### Concepts


[NavigationFolders Object](navigationfolders-object-outlook.md)

