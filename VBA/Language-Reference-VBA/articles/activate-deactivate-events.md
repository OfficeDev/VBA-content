---
title: Activate, Deactivate Events
keywords: vblr6.chm916300
f1_keywords:
- vblr6.chm916300
ms.prod: office
ms.assetid: 387d0954-5f02-9869-2709-35103634e7ae
ms.date: 06/08/2017
---


# Activate, Deactivate Events



The Activate event occurs when an object becomes the active window. The Deactivate event occurs when an object is no longer the active window.
 **Syntax**
 **Private Sub**_object_**_Activate()**
 **Private Sub**_object_**_Deactivate()**
The  _object_ placeholder represents an[object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.
 **Remarks**
An object can become active by using the  **Show** method in code.
The Activate event can occur only when an object is visible. A  **UserForm** loaded with **Load** isn't visible unless you use the **Show** method.
The Activate and Deactivate events occur only when you move the [focus](vbe-glossary.md) within an application. Moving the focus to or from an object in another application doesn't trigger either event.
The Deactivate event doesn't occur when unloading an object.

## Example

The following code uses two  **UserForms**: UserForm1 and UserForm2. Copy these procedures into the UserForm1 module, then add UserForm2. UserForm1's caption is created in its Activate event procedure. When the user clicks the client area of UserForm1, UserForm2 is loaded, and shown, triggering UserForm1's Deactivate event, changing their captions.


```vb
' Activate event for UserForm1
Private Sub UserForm_Activate()
    UserForm1.Caption = "Click my client area"
End Sub

' Click event for UserForm1
Private Sub UserForm_Click()
    Load UserForm2
    UserForm2.StartUpPosition = 3
    UserForm2.Show
End Sub

' Deactivate event for UserForm1
Private Sub UserForm_Deactivate()
    UserForm1.Caption = "I just lost the focus!"
    UserForm2.Caption = "Focus just left UserForm1 and came to me"
End Sub
```


