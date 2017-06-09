---
title: AppointmentItem.Display Method (Outlook)
keywords: vbaol11.chm871
f1_keywords:
- vbaol11.chm871
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Display
ms.assetid: 24706404-d646-a3ac-b7b1-64a6a1c697a9
ms.date: 06/08/2017
---


# AppointmentItem.Display Method (Outlook)

Displays a new  **[Inspector](inspector-object-outlook.md)** object for the item.


## Syntax

 _expression_ . **Display**( **_Modal_** )

 _expression_ A variable that represents an **AppointmentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Modal_|Optional| **Variant**| **True** to make the window modal. The default value is **False** .|

## Remarks

The  **Display** method is supported for explorer and inspector windows for the sake of backward compatibility. To activate an explorer or inspector window, use the **[Activate](inspector-activate-method-outlook.md)** method.

If you attempt to open an "unsafe" file system object (or "freedoc" file) by using the Microsoft Outlook object model, you receive the  **E_FAIL** return code in the C or C++ programming languages. In Outlook 2000 and earlier, you could open an "unsafe" file system object by using the **Display** method.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

