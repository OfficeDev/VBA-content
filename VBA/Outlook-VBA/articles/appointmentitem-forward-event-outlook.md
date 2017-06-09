---
title: AppointmentItem.Forward Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Forward
ms.assetid: 3d56ee04-9a9a-1f10-0436-a2b567b17229
ms.date: 06/08/2017
---


# AppointmentItem.Forward Event (Outlook)

Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **Forward**( **_Forward_** , **_Cancel_** )

 _expression_ A variable that represents an **AppointmentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Forward_|Required| **Object**|The new item being forwarded.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the forward operation is not completed and the new item is not displayed.|

## Remarks

In VBScript, if you set the return value of this function to  **False** , the forward action is not completed and the new item is not displayed.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

