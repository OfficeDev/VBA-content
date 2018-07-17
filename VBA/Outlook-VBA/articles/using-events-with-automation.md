---
title: Using Events with Automation
keywords: vbaol11.chm5257590
f1_keywords:
- vbaol11.chm5257590
ms.prod: outlook
ms.assetid: 6ca0a0fa-1cda-c052-4dee-1055cceb2b28
ms.date: 06/08/2017
---


# Using Events with Automation

To create an event handler for Microsoft Outlook objects in Microsoft Visual Basic or Microsoft Visual Basic for Applications (VBA) in another application, you need to complete the following four steps:


1. Set a reference to the Outlook Object Library.
    
2. Declare an object variable to respond to the events.
    
3. Write the specific event procedures.
    
4. Initialize the declared object.
    

Learn about  [working with events in Outlook Visual Basic for Applications](using-outlook-visual-basic-for-applications-to-respond-to-outlook-events.md).


## Set the Reference to the Outlook Object Library

Before you can use an Outlook object in Visual Basic or Visual Basic for Applications code, you must first set a reference to the Outlook Object Model in the  **References** dialog box. For more information about using this dialog box, see the online Help for your programming environment.


## Declare the Object Variable

Once you've referenced the object model library, you must declare variables that reference the object you want to use. You can declare the variable in the module in which the object will be used (that is, the module containing the event-handler procedure), but more commonly you' ll declare it in a class module so it can be used in any module in your program.

For example, to declare an object variable for the  **[Application](application-object-outlook.md)** object in a class module, you use code like the following.




```vb
Public WithEvents myOlApp As Outlook.Application
```

You must use the  `WithEvents` keyword to specify that the object variable will be used to respond to events triggered by the object.


## Write the Event Procedure

After the new object has been declared with events, it appears in the  **Object** list in the class module Code window, and you can select the object's event procedures from the **Procedures/Events** list. For example, when you select the **[ItemSend](application-itemsend-event-outlook.md)** event for an **Application** object declared as `myOlApp`, the following empty procedure appears in the Code window.


```vb
Private Sub myOlApp_ItemSend(Item as Object, Cancel as Boolean) 
 
End Sub
```


## Initialize the Declared Object

Before the procedure will run, you must connect the declared object (in this example,  `myOlApp`) with the  **Application** object. If you declared the object in a class module named `EventClassModule`, then you can use the following code in any module.


```vb
Dim myClass as New EventClassModule  
Sub Register_Event_Handler()  
    Set myClass.myOlApp = "Outlook.Application"  
End Sub
```

When the




```
Register_Event_Handler
```

procedure is run, the  `myOlApp` object in the form or class module points to the Outlook **Application** object, and the event procedure will run when the event occurs.


