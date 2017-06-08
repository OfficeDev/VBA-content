---
title: Label.Caption Property (Access)
keywords: vbaac10.chm10191
f1_keywords:
- vbaac10.chm10191
ms.prod: access
api_name:
- Access.Label.Caption
ms.assetid: 47f58d63-a93d-a0ef-333c-ab0479bad6c9
ms.date: 06/08/2017
---


# Label.Caption Property (Access)

Gets or sets the text that appears in the control. Read/write  **String**.


## Syntax

 _expression_. **Caption**

 _expression_ A variable that represents a **Label** object.


## Remarks

The  **Caption** property is a string expression that can contain up to 2,048 characters.

If you don't specify a caption for a table field, the field's  **FieldName** property setting will be used as the caption of a label attached to a control or as the column heading in Datasheet view. If you don't specify the caption for a query field, the caption for the underlying table field will be used. If you don't set a caption for a form, button, or label, Microsoft Access will assign the object a unique name based on the object, such as "Form1".

If you create a control by dragging a field from the field list and haven't specified a  **Caption** property setting for the field, the field's **FieldName** property setting will be copied to the control's **Name** property box and will also appear in the label of the control created.


 **Note**  The text of the  **Caption** property for a label or command button control is the hyperlink display text when the **HyperlinkAddress** or **HyperlinkSubAddress** property is set for the control.

You can use the  **Caption** property to assign an access key to a label or command button. In the caption, include an ampersand (&;) immediately preceding the character you want to use as an access key. The character will be underlined. You can press ALT plus the underlined character to move the focus to that control on a form.

Include two ampersands (&;&;) in the setting for a caption if you want to display an ampersand itself in the caption text. For example, to display "Save &; Exit", you should type  **Save &;&; Exit** in the **Caption** property box.


## See also


#### Concepts


[Label Object](label-object-access.md)

