---
title: Form.CurrentSectionLeft Property (Access)
keywords: vbaac10.chm13468
f1_keywords:
- vbaac10.chm13468
ms.prod: access
api_name:
- Access.Form.CurrentSectionLeft
ms.assetid: 5c856f2a-f82c-2b67-6fc6-1773fc5ebe06
ms.date: 06/08/2017
---


# Form.CurrentSectionLeft Property (Access)

You can use this property to determine the distance in twips from the left side of the current section to the left side of the form. Read/write  **Integer**.


## Syntax

 _expression_. **CurrentSectionLeft**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **CurrentSectionLeft** property setting changes whenever a user scrolls through a form.

For forms whose  **[DefaultView](form-defaultview-property-access.md)** property is set to Single Form, if the user scrolls to the right of the left edge of the form, the property setting is a negative value.

The  **CurrentSectionLeft** property is useful for finding the positions of detail sections displayed in Form view as continuous forms or in Datasheet view.


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

