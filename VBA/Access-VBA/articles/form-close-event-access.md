---
title: Form.Close Event (Access)
keywords: vbaac10.chm13645
f1_keywords:
- vbaac10.chm13645
ms.prod: access
api_name:
- Access.Form.Close
ms.assetid: e65fe7e0-efc1-dabc-4b2c-787af465ade0
ms.date: 06/08/2017
---


# Form.Close Event (Access)

The  **Close** event occurs when a form is closed and removed from the screen.


## Syntax

 _expression_. **Close**

 _expression_ A variable that represents a **Form** object.


### Return Value

nothing


## Remarks

To run a macro or event procedure when this event occurs, set the  **OnClose** property to the name of the macro or to [Event Procedure].

The  **Close** event occurs after the **Unload** event, which is triggered after the form is closed but before it is removed from the screen.

When you close a form, the following events occur in this order:

 **Unload** → **Deactivate** → **Close**

When the  **Close** event occurs, you can open another window or request the user's name to make a log entry indicating who used the form or report.

The  **Unload** event can be canceled, but the **Close** event can't.


## See also


#### Concepts


[Form Object](form-object-access.md)

