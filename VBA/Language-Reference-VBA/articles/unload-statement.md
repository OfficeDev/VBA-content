---
title: Unload Statement
keywords: vblr6.chm1100684
f1_keywords:
- vblr6.chm1100684
ms.prod: office
ms.assetid: 5fa03dfb-686d-b266-18ba-e4c50afd63ea
ms.date: 06/08/2017
---


# Unload Statement

Removes an object from memory.

 **Syntax**

 **Unload**_object_

The required  _object_ placeholder represents an[object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.
 **Remarks**
When an object is unloaded, it's removed from memory and all memory associated with the object is reclaimed. Until it is placed in memory again using the  **Load** statement, a user can't interact with an object, and the object can't be manipulated programmatically.

## Example

The following example assumes two  **UserForms** in a program. In UserForm1's Initialize event, UserForm2 is loaded and shown. When the user clicks UserForm2, it is unloaded and UserForm1 appears. When UserForm1 is clicked, it is unloaded in turn.


```vb
' This is the Initialize event procedure for UserForm1 
Private Sub UserForm_Initialize() 
 Load UserForm2 
 UserForm2.Show 
End Sub 
' This is the Click event for UserForm2 
Private Sub UserForm_Click() 
 Unload UserForm2 
End Sub 
 
' This is the Click event for UserForm1 
Private Sub UserForm_Click() 
 Unload UserForm1 
End Sub
```


