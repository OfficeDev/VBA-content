---
title: DoCmd.CancelEvent Method (Access)
keywords: vbaac10.chm4144
f1_keywords:
- vbaac10.chm4144
ms.prod: access
api_name:
- Access.DoCmd.CancelEvent
ms.assetid: f8c0d2ff-9bf3-09d5-d15b-d3134bb6df80
ms.date: 06/08/2017
---


# DoCmd.CancelEvent Method (Access)

The  **CancelEvent** method carries out the CancelEvent action in Visual Basic.


## Syntax

 _expression_. **CancelEvent**

 _expression_ A variable that represents a **DoCmd** object.


## Remarks

You can use the CancelEvent method to cancel the event that caused Microsoft Access to run the procedure containing this method. 

The  **CancelEvent** method has an effect only when it's run as the result of an event. This method cancels the event.

In a form, you typically use the CancelEvent action in a validation macro with the  **BeforeUpdate** event property. When a user enters data in a control or record, Access runs the macro before adding the data to the database. If the data fails the validation conditions in the macro, the CancelEvent action cancels the update process before it starts.

All events that can be canceled in Visual Basic have a Cancel argument. You can use this argument instead of the  **CancelEvent** method to cancel the event. The **KeyPress** event and **MouseDown** event (for right-clicking only) can be canceled only in macros, not event procedures, so you must use the CancelEvent action in a macro to cancel these events.


 **Note**  You can use the  **CancelEvent** method with the **MouseDown** event only to cancel the event that occurs when you right-click an object.

For events that can be canceled, the default behavior for the event (that is, what Access typically does when the event occurs) occurs after the procedure for the event runs. This enables you to cancel the default behavior. For example, when you double-click a word that the insertion point is on in a text box, Access normally selects the word. You can cancel this default behavior in the procedure for the  **DblClick** event and perform some other action, such as opening a form containing information about the data in the text box. For events that can't be canceled, the default behavior occurs before the procedure runs.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

