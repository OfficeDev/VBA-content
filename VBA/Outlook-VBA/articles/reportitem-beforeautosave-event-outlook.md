---
title: ReportItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ReportItem.BeforeAutoSave
ms.assetid: c3a2882c-ff82-39a1-3d18-5bf4f608b09e
ms.date: 06/08/2017
---


# ReportItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** )

 _expression_ A variable that represents a **ReportItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[ReportItem](reportitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

