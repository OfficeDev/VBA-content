---
title: DistListItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DistListItem.BeforeAutoSave
ms.assetid: bb005bda-6270-22a8-5ae0-43979e3f3e63
ms.date: 06/08/2017
---


# DistListItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_. **BeforeAutoSave** (**_Cancel_**)

 _expression_ A variable that represents a **DistListItem** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Boolean**|Set to **True** to cancel the operation; otherwise, set to **False** to allow the **[DistListItem](distlistitem-object-outlook.md)** to be saved.|

<br/>

## See also

#### Concepts

- [DistListItem Object](distlistitem-object-outlook.md)

