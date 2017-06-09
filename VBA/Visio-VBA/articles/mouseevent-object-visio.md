---
title: MouseEvent Object (Visio)
keywords: vis_sdr.chm52060
f1_keywords:
- vis_sdr.chm52060
ms.prod: visio
api_name:
- Visio.MouseEvent
ms.assetid: 1ae26c28-8fdd-ecfe-b008-d4788c08ce5a
ms.date: 06/08/2017
---


# MouseEvent Object (Visio)

The object passed to  **VisEventProc** as the subject of **MouseDown** , **MouseMove** , and **MouseUp** events.


## Remarks

The default property of  **MouseEvent** is **ToString** . The **ToString** property returns a string that represents the properties of the **MouseEvent** object and has the form

 _event code_;  **Button** property value; **KeyButtonState** property value; **x** property value; **y** property value; **Window.Caption**

where  _event code_ returns the code of the event that fired ( **MouseDown** , **MouseMove** , or **MouseUp** ) and **Window.Caption** returns the caption of the window that sourced the event. For example, if a user clicked the left mouse button near the middle of the drawing page while holding down the SHIFT key, in response to the **MouseDown** event, **ToString** might return

709;1;5;4.3750003+000;4.265000+000;Drawing1

Use the  **Application** property of the **MouseEvent** object to determine the Microsoft Visio instance hosting the object, and use the **Window** property to determine the Visio window associated with a mouse event.


