---
title: Hide Method
keywords: vblr6.chm916117
f1_keywords:
- vblr6.chm916117
ms.prod: office
api_name:
- Office.Hide
ms.assetid: 24844c21-0181-24e9-10f6-2ac006f99cbe
ms.date: 06/08/2017
---


# Hide Method



Hides an object but doesn't unload it.
 **Syntax**
 _object_. **Hide**
The  _object_ placeholder represents an[object expression](vbe-glossary.md) that evaluates to an object in the Applies To list. If _object_ is omitted, the **UserForm** with the[focus](vbe-glossary.md) is assumed to be _object_.
 **Remarks**
When an object is hidden, it's removed from the screen and its  **Visible** property is set to **False**. A hidden object's controls aren't accessible to the user, but they are available programmatically to the running application, to other processes that may be communicating with the application through Automation, and in Windows, to **Timer** control events.
When a  **UserForm** is hidden, the user can't interact with the application until all code in the event procedure that caused the **UserForm** to be hidden has finished executing.
If the  **UserForm** isn't loaded when the **Hide** method is invoked, the **Hide** method loads the **UserForm** but doesn't display it.

## Example

The following example assumes two  **UserForms** in a program. In UserForm1's Initialize event, UserForm2 is loaded and shown. When the user clicks UserForm2, it is hidden and UserForm1 appears. When UserForm1 is clicked, UserForm2 is shown again.


```vb
' This is the Initialize event procedure for UserForm1
Private Sub UserForm_Initialize()
    Load UserForm2
    UserForm2.Show
End Sub
' This is the Click event of UserForm2
Private Sub UserForm_Click()
    UserForm2.Hide
End Sub

' This is the click event for UserForm1
Private Sub UserForm_Click()
    UserForm2.Show
End Sub
```


