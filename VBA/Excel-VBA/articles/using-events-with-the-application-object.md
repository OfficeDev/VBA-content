---
title: Using Events with the Application Object
keywords: vbaxl10.chm5205784
f1_keywords:
- vbaxl10.chm5205784
ms.prod: excel
ms.assetid: 0063feba-47fd-29be-d2d5-8fcf47e70cbc
ms.date: 06/08/2017
---


# Using Events with the Application Object

Before you can use events with the  **Application** object, you must create a class module and declare an object of type **Application** with events. For example, assume that a new class module is created and called EventClassModule. The new class module contains the following code:


```vb
Public WithEvents App As Application
```


After the new object has been declared with events, it appears in the  **Object** drop-down list box in the class module, and you can write event procedures for the new object. (When you select the new object in the **Object** box, the valid events for that object are listed in the **Procedure** drop-down list box.)

Before the procedures will run, however, you must connect the declared object in the class module with the  **Application** object. You can do this with the following code from any module.

## Example


```vb
Dim X As New EventClassModule 
 
Sub InitializeApp() 
 Set X.App = Application 
End Sub
```

After you run the  **InitializeApp** procedure, the **App** object in the class module points to the Microsoft Excel **Application** object, and the event procedures in the class module will run when the events occur.


