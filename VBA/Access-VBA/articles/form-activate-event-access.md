---
title: Form.Activate Event (Access)
keywords: vbaac10.chm13646
f1_keywords:
- vbaac10.chm13646
ms.prod: access
api_name:
- Access.Form.Activate
ms.assetid: 1409c52b-8a77-0e0d-1a26-7dc4ce8bb320
ms.date: 06/08/2017
---


# Form.Activate Event (Access)

The Activate event occurs when a form receives the focus and becomes the active window.


## Syntax

 _expression_. **Activate**

 _expression_ A variable that represents a **Form** object.


### Return Value

nothing


## Remarks


 **Note**  The  **Activate** event doesn't occur when a form receives focus back from a dialog box, popup, or another form.

To run a macro or event procedure when these events occur, set the  **OnActivate**, or **OnDeactivate** property to the name of the macro or to [Event Procedure].

You can make a form active by opening it, clicking it or a control on it, or by using the  **SetFocus** method in Visual Basic.

The  **Activate** event can occur only when a form is visible.

The  **Activate** event occurs before the **GotFocus** event; the Deactivate event occurs after the **LostFocus** event.

When you switch between two open forms, the  **Deactivate** event occurs for the form being switched from, and the **Activate** event occurs for the form being switched to. If the forms contain no visible, enabled controls, the **LostFocus** event occurs for the first form before the **Deactivate** event, and the **GotFocus** event occurs for the second form after the **Activate** event.

When you first open a form, the following events occur in this order:

 **Open** → **Load** → **Resize** → **Activate** → **Current**

When you close a form, the following events occur in this order:

 **Unload** → **Deactivate** → **Close**


## See also


#### Concepts


[Form Object](form-object-access.md)

