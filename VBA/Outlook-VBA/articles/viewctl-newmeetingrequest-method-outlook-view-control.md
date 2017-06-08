---
title: ViewCtl.NewMeetingRequest Method (Outlook View Control)
ms.prod: outlook
ms.assetid: b8c76fcf-e44c-94a1-2ada-0347c14b70cf
ms.date: 06/08/2017
---


# ViewCtl.NewMeetingRequest Method (Outlook View Control)

Creates and displays a new meeting request.


## Syntax

 _expression_. **NewMeetingRequest**

 _expression_A variable that represents a  **ViewCtl** object.


## Remarks

When the meeting request is sent, the corresponding appointment is saved to the  **Calendar**folder, if any, that is displayed in the control. If there is no folder displayed in the control, the appointment is saved to the user's default  **Calendar** folder.

Responses to the meeting request are tallied only if the appointment is saved to the user's default  **Calendar** folder.


