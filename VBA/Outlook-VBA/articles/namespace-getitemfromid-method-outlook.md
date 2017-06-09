---
title: NameSpace.GetItemFromID Method (Outlook)
keywords: vbaol11.chm763
f1_keywords:
- vbaol11.chm763
ms.prod: outlook
api_name:
- Outlook.NameSpace.GetItemFromID
ms.assetid: f2abff80-4c04-998b-654b-28600424a16f
ms.date: 06/08/2017
---


# NameSpace.GetItemFromID Method (Outlook)

Returns a Microsoft Outlook item identified by the specified entry ID (if valid). 


## Syntax

 _expression_ . **GetItemFromID**( **_EntryIDItem_** , **_EntryIDStore_** )

 _expression_ A variable that represents a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EntryIDItem_|Required| **String**| The **[EntryID](folder-entryid-property-outlook.md)** of the item.|
| _EntryIDStore_|Optional| **Variant**|The  **[StoreID](folder-storeid-property-outlook.md)** for the folder. _EntryIDStore_ usually must be provided when retrieving an item based on its MAPI IDs.|

### Return Value

An  **Object** value that represents the specified Outlook item.


## Remarks

This method is used for ease of transition between MAPI and OLE/Messaging applications and Outlook.

For more information about Entry IDs, see the  **[EntryID](folder-entryid-property-outlook.md)** property.


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

