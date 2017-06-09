---
title: AppointmentItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.BeforeAutoSave
ms.assetid: c24e39d1-39e5-6422-78ff-9d4e391ea2ae
ms.date: 06/08/2017
---


# AppointmentItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** , )

 _expression_ A variable that represents an **AppointmentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[AppointmentItem](appointmentitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

