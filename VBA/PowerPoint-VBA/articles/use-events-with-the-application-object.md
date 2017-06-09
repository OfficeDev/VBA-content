---
title: Use Events with the Application Object
keywords: vbapp10.chm5237994
f1_keywords:
- vbapp10.chm5237994
ms.prod: powerpoint
ms.assetid: b657ab62-67fa-4eeb-736c-86e31a026c73
ms.date: 06/08/2017
---


# Use Events with the Application Object

To create an event handler for an event of the  **Application** object, you need to complete the following three steps:


1. Declare an object variable in a class module to respond to the events.
    
2. Write the specific event procedures.
    
3. Initialize the declared object from another module.
    

## Declare the Object Variable

Before you can write procedures for the events of the  **Application** object, you must create a new class module and declare an object of type **Application** with events. For example, assume that a new class module is created and called EventClassModule. The new class module contains the following code.


```vb
Public WithEvents App As Application
```


## Write the Event Procedures

After the new object is declared with events, it appears in the  **Object** list in the class module, and you can write event procedures for the new object. (When you select the new object in the **Object** list, the valid events for that object are listed in the **Procedure** list.) Select an event from the **Procedure** list; an empty procedure is added to the class module.


```vb
Private Sub App_NewPresentation()

End Sub
```


## Initializing the Declared Object

Before the procedure will run, you must connect the declared object in the class module (App in this example) with the  **Application** object. You can do this with the following code from any module.


```vb
Dim X As New EventClassModule
Sub InitializeApp()
    Set X.App = Application
End Sub
```

Run the InitializeApp procedure. After the procedure is run, the App object in the class module points to the Microsoft Office PowerPoint  **Application** object, and the event procedures in the class module will run when the events occur.


