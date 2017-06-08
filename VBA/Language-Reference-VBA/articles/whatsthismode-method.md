---
title: WhatsThisMode Method
keywords: vblr6.chm1100685
f1_keywords:
- vblr6.chm1100685
ms.prod: office
api_name:
- Office.WhatsThisMode
ms.assetid: e71fb00c-b323-2b43-94ec-07079e66337f
ms.date: 06/08/2017
---


# WhatsThisMode Method



Causes the mouse pointer to change to the  **What's This** pointer and prepares the application to display Help on a selected object. This method exists on the Macintosh, but there is no pointer functionality.
 **Syntax**
 _object_. **WhatsThisMode**
The  _object_ placeholder represents an[object expression](vbe-glossary.md) that evaluates to an object in the Applies To list. If _object_ is omitted, the **UserForm** with the[focus](vbe-glossary.md) is assumed to be _object_.
 **Remarks**
Executing the  **WhatsThisMode** method places the application in the same state as clicking the **What's This** button on the title bar. The mouse pointer changes to the **What's This** pointer. When the user clicks an object, the **WhatsThisHelpID** property of the clicked object is used to invoke the context-sensitive Help.

## Example

The following example changes the mouse pointer to the  **What's This** (question mark) pointer when the user clicks the **UserForm**. If neither the **WhatsThisHelp** or the **WhatsThisButton** property is set to **True** in the **Properties** window, the following invocation has no effect.


```vb
Private Sub UserForm_Click()
' Turn mouse pointer to What's This question mark.
    WhatsThisMode
End Sub
```


