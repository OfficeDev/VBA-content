---
title: Pages.ItemFromID Property (Visio)
keywords: vis_sdr.chm11013775
f1_keywords:
- vis_sdr.chm11013775
ms.prod: visio
api_name:
- Visio.Pages.ItemFromID
ms.assetid: 0355a186-b7bf-51e5-bb2c-433417cf2d33
ms.date: 06/08/2017
---


# Pages.ItemFromID Property (Visio)

Returns an item of a collection using the ID of the item. Read-only.


## Syntax

 _expression_ . **ItemFromID**( **_nID_** )

 _expression_ A variable that represents a **Pages** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _nID_|Required| **Long**|The ID of the object to retrieve.|

### Return Value

Page


## Remarks

The ID of a  **Shape** object uniquely identifies the shape within its page or master.

The ID of a  **Style** object uniquely identifies the style within its document.

The ID of a  **Font** object corresponds to the number stored in the Font cell of a row in a shape's Character Properties section. The ID associated with a particular font varies between systems or as fonts are installed on and removed from a given system.

The ID of an  **Event** object uniquely identifies an event in its **EventList** collection for the life of the collection.


