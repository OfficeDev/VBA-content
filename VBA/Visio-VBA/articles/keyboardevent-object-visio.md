---
title: KeyboardEvent Object (Visio)
keywords: vis_sdr.chm52055
f1_keywords:
- vis_sdr.chm52055
ms.prod: visio
api_name:
- Visio.KeyboardEvent
ms.assetid: 5091c972-b226-1caa-d40f-96a5f3b5bf01
ms.date: 06/08/2017
---


# KeyboardEvent Object (Visio)

The object passed to  **VisEventProc** as the subject of **KeyDown** , **KeyPress** , and **KeyUp** events.


## Remarks

The default property of  **KeyboardEvent** is **ToString** . The **ToString** property returns a string that represents the properties of the **KeyboardEvent** object, and has the form

 _event code_ ; **KeyCode** property value; **KeyButtonState** property value; **KeyAscii** property value; **Window.Caption**

where  _event code_ returns the code of the event that fired ( **KeyDown** , **KeyPress** , or **KeyUp** ) and **Window.Caption** returns the caption of the window that sourced the event. For example, if a user pressed the "L" key while holding down the SHIFT key, in response to the **KeyPress** event, **ToString** might return the following:

713;0;4;76;Drawing1

Use the  **Application** property of the **KeyboardEvent** object to determine the Microsoft Visio instance hosting the object, and use the **Window** property to determine the Visio window associated with a keyboard event.


