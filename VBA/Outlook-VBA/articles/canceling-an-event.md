---
title: Canceling an Event
keywords: olfm10.chm3077110
f1_keywords:
- olfm10.chm3077110
ms.prod: outlook
ms.assetid: ee23d8d9-d815-f09e-d87a-dd2db71ef093
ms.date: 06/08/2017
---


# Canceling an Event



 Outlook calls event handlers in your program to allow your program to respond to such events as actions that the user takes or changes in the message store. Each event is accompanied by a default action that Outlook performs as a result of the event. For example, when the **Open** event occurs for an item, by default Outlook displays the item in an inspector window.

Some events only notify your program that a particular event has occurred. For these events, your event handler simply responds to the event. With other events, Outlook allows your event handler to cancel the event, that is, to instruct Outlook not to perform the default action associated with the event. In the case of the  **Open** event, for example, your program can prevent Outlook from displaying the item in an inspector. If an event can be cancelled, the reference topic describing the event indicates how to cancel the event.

If an event can be cancelled, an event handler written in Microsoft Visual Basic or Microsoft Visual Basic for Applications receives a parameter that it sets before returning to indicate whether the event should be cancelled. For example, an event handler for the  **Open** event written in Visual Basic for Applications might look like this. This example assumes that the value of OpenOK is set elsewhere.



```vb
Sub myItem_Open(byRef Cancel as Boolean) 
 If OpenOK Then 
 Cancel = False ' Outlook performs default action 
 Else 
 Cancel = True ' Outlook does not perform default action 
 EndIf 
End Sub
```

Because of limitations in VBScript, however, this syntax cannot be used. An event handler for the  **Open** event in the script of an item must be written as a function. To cancel the event, the value of the function is set to **False** before returning, as in the following example.



```vb
Function Item_Open() 
 If OpenOK Then 
 Item_Open = True ' Outlook performs default action 
 Else 
 Item_Open = False ' Outlook does not perform default action 
 End If 
End Function
```


