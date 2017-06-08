---
title: Form.CurrentSectionTop Property (Access)
keywords: vbaac10.chm13467
f1_keywords:
- vbaac10.chm13467
ms.prod: access
api_name:
- Access.Form.CurrentSectionTop
ms.assetid: d6f4f5f6-641f-3092-7d99-195c77722718
ms.date: 06/08/2017
---


# Form.CurrentSectionTop Property (Access)

You can use this property to determine the distance in twips from the top edge of the current section to the top edge of the form. Read/write  **Integer**.


## Syntax

 _expression_. **CurrentSectionTop**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **CurrentSectionTop** property setting changes whenever a user scrolls through a form.

For forms whose  **[DefaultView](form-defaultview-property-access.md)** property is set to Single Form, if the user scrolls above the upper-left corner of the section, the property settings are negative values.

For forms whose  **DefaultView** property is set to Continuous Forms, if a section isn't visible, the **CurrentSectionTop** property is equal to the **[InsideHeight](form-insideheight-property-access.md)** property of the form.

The  **CurrentSectionTop** property is useful for finding the positions of detail sections displayed in Form view as continuous forms or in Datasheet view. Each detail section has a different **CurrentSectionTop** property setting, depending on the section's position on the form.


## Example

The following example displays the  **CurrentSectionLeft** and **CurrentSectionTop** property settings for a control on a continuous form. Whenever the user moves to a new record, the property settings for the current section are displayed in the `lblStatus` label in the form's header.


```vb
Private Sub Form_Current() 
 
 Dim intCurTop As Integer 
 Dim intCurLeft As Integer 
 
 intCurTop = Me.CurrentSectionTop 
 intCurLeft = Me.CurrentSectionLeft 
 Me!lblStatus.Caption = intCurLeft &; " , " &; intCurTop 
 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

