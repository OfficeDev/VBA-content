---
title: Using Outlook Visual Basic for Applications to Respond to Outlook Events
keywords: vbaol11.chm5274250
f1_keywords:
- vbaol11.chm5274250
ms.prod: outlook
ms.assetid: 560bb264-05d0-dbc6-39c2-b95b12f50ed9
ms.date: 06/08/2017
---


# Using Outlook Visual Basic for Applications to Respond to Outlook Events

You write an event procedure (also known as an event handler) to respond to events that occur in Microsoft Outlook. For example, you can write an event procedure that automatically maximizes the explorer window when Outlook starts.

Events are associated with particular objects. The  [Application](application-object-outlook.md) object is the topmost object, and is always available (that is, it does not have to be created). You can add an **Application** event procedure in the **ThisOutlookSession** module window simply by selecting **Application** in the left list and then selecting the event in the right list.

Adding an event handler for objects other than the  **Application** object requires a few additional steps.

First, you must declare a variable using the  **WithEvents** keyword to identify the object whose event you want to handle. For example, to declare a variable representing the [OutlookBarPane](outlookbarpane-object-outlook.md) object, you would add the following to a code module.



```vb
Dim WithEvents myOlBar as Outlook.OutlookBarPane
```

You can then select  `myOlBar` in the Objects list of the module window and then select the event in the procedure list. The Visual Basic Editor will then add the template for the event procedure to the module window. You can then type the code you want to run when the event occurs. The following example shows code added to the [BeforeNavigate](outlookbarpane-beforenavigate-event-outlook.md) event procedure for the **OutlookBarPane** object.



```vb
Private Sub myOlBar_BeforeNavigate(ByVal Shortcut As OutlookBarShortcut, Cancel As Boolean) 
 If Shortcut.Name = "Notes" Then 
 MsgBox "You cannot open the Notes folder." 
 Cancel = True 
 End If 
End Sub
```

The final step is to add code to set the object variable to the object whose event you want to handle. This code can exist in a macro, or if you want the event to be handled whenever Outlook runs, you can put it in the  [Startup](application-startup-event-outlook.md) event procedure, as in the following example.



```vb
Private Sub Application_Startup() 
 Set myOlBar = Application.ActiveExplorer.Panes(1) 
End Sub
```


