---
title: MasterShortcuts.ItemFromID Property (Visio)
keywords: vis_sdr.chm15913775
f1_keywords:
- vis_sdr.chm15913775
ms.prod: visio
api_name:
- Visio.MasterShortcuts.ItemFromID
ms.assetid: 4a4eb2f9-9ab7-4574-c942-e9bd4031fabd
ms.date: 06/08/2017
---


# MasterShortcuts.ItemFromID Property (Visio)

Returns an item of a collection using the ID of the item. Read-only.


## Syntax

 _expression_ . **ItemFromID**( **_nID_** )

 _expression_ A variable that represents a **MasterShortcuts** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _nID_|Required| **Long**|The ID of the object to retrieve.|

### Return Value

MasterShortcut


## Remarks

The ID of a  **Shape** object uniquely identifies the shape within its page or master.

The ID of a  **Style** object uniquely identifies the style within its document.

The ID of a  **Font** object corresponds to the number stored in the Font cell of a row in a shape's Character Properties section. The ID associated with a particular font varies between systems or as fonts are installed on and removed from a given system.

The ID of an  **Event** object uniquely identifies an event in its **EventList** collection for the life of the collection.


