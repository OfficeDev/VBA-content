---
title: AppointmentItem.Send Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Send
ms.assetid: 6571ae2f-4964-f38f-e39e-14a2b94caa73
ms.date: 06/08/2017
---


# AppointmentItem.Send Event (Outlook)

Occurs when the user selects the  **Send** action for an item, or when the **Send** method is called for the item, which is an instance of the parent object.


## Syntax

 _expression_ . **Send**( **_Cancel_** )

 _expression_ A variable that represents an **AppointmentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the item is not sent.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

