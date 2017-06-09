---
title: Working with Outlook Events
keywords: vbaol11.chm5267482
f1_keywords:
- vbaol11.chm5267482
ms.prod: outlook
ms.assetid: 514f8f31-8047-2a9f-cbac-d0a23218f49c
ms.date: 06/08/2017
---


# Working with Outlook Events

 Outlook provides a wide range of events through which it can notify your Microsoft Visual Basic, Microsoft Visual Basic for Applications (VBA), and Microsoft Visual Basic Scripting Edition (VBScript) programs that a significant change has occurred. For example, Outlook events can notify a program when an item has been opened or when a new mail arrives in the InBox.

To receive notification of a significant event, write an event-handler procedure. Depending on whether the event is handled in Visual Basic or Visual Basic for Applications or in VBScript, this is either a  `Sub` or a `Function` that Outlook calls when the event is called. The code you put in the event handler allows your program to respond appropriately to the event and, in some cases, even lets your program cancel the default action associated with the event, such as preventing a mail item from being sent.

## Types of Events

Outlook events can be divided into two main categories: item-level events and application-level events.

Item-level events pertain to a particular item, and are typically handled by VBScript code contained within the form associated with the item. These events notify your program when an item has been opened, sent or posted, saved, or closed, and when the user has replied to or forwarded a message or initiated a custom action. Item-level events can also notify your program when the user has clicked a control on the form or when an item property has changed.

Application-level events are typically handled by Visual Basic or Visual Basic for Applications because they pertain to more than the items associated with a particular form. Application-level events can pertain to the application itself, to explorer collections and windows (including the Shortcuts pane), inspector collections and windows, folders and folders collections, items collections, and synchronization objects.


## Responding to Events

To respond to item-level events, add event-handler procedures to the script of the form that displays the item. For example, to run code when an item is opened in the form, add a procedure like the following to the script in the form.


```vb
Function Item_Open() 
 MsgBox "A new item has opened in this form." 
End Function
```

Responding to application-level events is somewhat more involved because steps must be taken to associate the event handler with the part of Outlook in which the event is occurring. Learn about  [writing an application-level event handler](using-events-with-automation.md).


## Order of Events

Except for certain form events, your program cannot assume that events will occur in a particular order, even if they appear to be called in a consistent sequence. The order in which Outlook calls event handlers might change depending on other events that might occur, or the order might change in future versions of Outlook.


