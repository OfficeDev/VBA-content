---
title: DropButtonClick Event
keywords: fm20.chm2000090
f1_keywords:
- fm20.chm2000090
ms.prod: office
api_name:
- Office.DropButtonClick
ms.assetid: 228f625c-937d-13ef-e04d-0d49a55fc0fd
ms.date: 06/08/2017
---


# DropButtonClick Event



Occurs whenever the drop-down list appears or disappears.
 **Syntax**
 **Private Sub**_object_ _**DropButtonClick( )**
The  **DropButtonClick** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
You can initiate the DropButtonClick event through code or by taking certain actions in the user interface.
In code, calling the  **DropDown** method initiates the DropButtonClick event.
In the user interface, any of the following actions initiates the event:


- Clicking the drop-down button on the control.
    
- Pressing F4.
    

Any of the above actions, in code or in the interface, cause the drop-down box to appear on the control. The system initiates the DropButtonClick event when the drop-down box goes away.

