---
title: AppointmentItem.Write Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Write
ms.assetid: 55539ad2-d53e-b28e-06f4-13c5f545a89b
ms.date: 06/08/2017
---


# AppointmentItem.Write Event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](appointmentitem-save-method-outlook.md)** or **[SaveAs](appointmentitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

 _expression_ . **Write**( **_Cancel_** )

 _expression_ A variable that represents an **AppointmentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True** , the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the save operation is not completed.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

