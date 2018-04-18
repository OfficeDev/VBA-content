---
title: Show Method
keywords: vblr6.chm916142
f1_keywords:
- vblr6.chm916142
ms.prod: office
api_name:
- Office.Show
ms.assetid: 25d05f62-8901-5592-00e0-ca7c340dfb86
ms.date: 06/08/2017
---


# Show Method



Displays a  **UserForm** object.
 **Syntax**
[ _object_**.** ] **Show**_modal_
The  **Show** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Optional. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list. If _object_ is omitted, the **UserForm** associated with the active **UserForm**[module](vbe-glossary.md) is assumed to be _object._|
| _modal_|Optional. Variant value that determines if the  **UserForm** is modal or modeless.|
 **Settings**
The settings for  _modal_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbModal**|1|**UserForm** is modal. Default.|
|**vbModeless**|0|**UserForm** is modeless.|
 **Remarks**
If the specified object isn't loaded when the  **Show** method is invoked, Visual Basic automatically loads it.

 **Note**  In Microsoft Office 97, if a  **UserForm** is set to display as modeless, it causes a run-time error; Office 97 **UserForms** are always modal.

When a  **UserForm** is modeless, subsequent code is executed as it's encountered. Modeless forms do not appear in the task bar and are not in the window tab order.

 **Note**  You may lose data associated with a modeless  **UserForm** if you make a change to the **UserForm** project that causes it to recompile, for example, removing a code module.

When a  **UserForm** is modal, the user must respond before using any other part of the application. No subsequent code is executed until the **UserForm** is hidden or unloaded. Although other forms in the application are disabled when a **UserForm** is displayed, other applications are not.

## Example

The following example assumes two  **UserForms** in a program. In UserForm1's Initialize event, UserForm2 is loaded and shown. When the user clicks UserForm2, it is hidden and UserForm1 appears. When UserForm1 is clicked, UserForm2 is shown again.


```vb
' This is the Initialize event procedure for UserForm1
Private Sub UserForm_Initialize()
    Load UserForm2
    UserForm2.Show
End Sub
' This is the Click event for UserForm2
Private Sub UserForm_Click()
    UserForm2.Hide
End Sub

' This is the click event for UserForm1
Private Sub UserForm_Click()
    UserForm2.Show
End Sub
```


