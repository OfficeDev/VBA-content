---
title: PrintForm Method
keywords: vblr6.chm916130
f1_keywords:
- vblr6.chm916130
ms.prod: office
api_name:
- Office.PrintForm
ms.assetid: d4481074-6ecf-b845-2a51-ef34dcdc82ab
ms.date: 06/08/2017
---


# PrintForm Method



Sends a bit-by-bit image of a  **UserForm** object to the printer.
 **Syntax**
 _object_**.PrintForm**
The  _object_ placeholder represents an[object expression](vbe-glossary.md) that evaluates to an object in the Applies To list. If _object_ is omitted, the **UserForm** with the[focus](vbe-glossary.md) is assumed to be _object_.
 **Remarks**
 **PrintForm** prints all visible objects and[bitmaps](vbe-glossary.md) of the **UserForm** object. **PrintForm** also prints graphics added to a **UserForm** object.
The printer used by  **PrintForm** is determined by the operating system's **Control Panel** settings.

## Example

In the following example, the client area of the form is printed when the user clicks the form.


```vb
' This is the click event for UserForm1
Private Sub UserForm_Click()
    UserForm1.PrintForm
End Sub
```


