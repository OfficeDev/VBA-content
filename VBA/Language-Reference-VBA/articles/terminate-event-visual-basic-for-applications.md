---
title: Terminate Event (Visual Basic for Applications)
keywords: vblr6.chm1107499
f1_keywords:
- vblr6.chm1107499
ms.prod: office
ms.assetid: f386e522-fc8a-f073-668d-e804dca9de49
ms.date: 06/08/2017
---


# Terminate Event (Visual Basic for Applications)



Occurs when all references to an instance of an object are removed from memory by setting all [variables](vbe-glossary.md) that refer to the object to **Nothing** or when the last reference to the object goes out of[scope](vbe-glossary.md).
 **Syntax**
 **Private Sub**_object_**_Terminate( )**
The  _object_ placeholder represents an[object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.
 **Remarks**
The Terminate event occurs after the object is unloaded. The Terminate event isn't triggered if the instances of the  **UserForm** or[class](vbe-glossary.md) are removed from memory because the application terminated abnormally. For example, if your application invokes the **End** statement before removing all existing instances of the class or **UserForm** from memory, the Terminate event isn't triggered for that class or **UserForm**.



```vb
Private Sub UserForm_Activate()
    UserForm1.Caption = "Click me to kill me!"
End Sub

Private Sub UserForm_Click()
  Unload Me
End Sub

Private Sub UserForm_Terminate()
    Dim Count As Integer
    For Count = 1 To 100
        Beep
    Next
End Sub
```


## Example

The following event procedures cause a  **UserForm** to beep for a few seconds after the user clicks the client area to dismiss the form.


