---
title: OptionGroup.TabIndex Property (Access)
keywords: vbaac10.chm10832
f1_keywords:
- vbaac10.chm10832
ms.prod: access
api_name:
- Access.OptionGroup.TabIndex
ms.assetid: 1ceeb9cd-e9b6-129f-72b9-3d15d9622709
ms.date: 06/08/2017
---


# OptionGroup.TabIndex Property (Access)

You can use the  **TabIndex** property to specify a control's place in the tab order on a form or report. Read/write **Integer**.


## Syntax

 _expression_. **TabIndex**

 _expression_ A variable that represents an **OptionGroup** object.


## Remarks

You can set the  **TabIndex** property to an integer representing the position of the control within the tab order of the form. Valid settings are 0 for the first tab position, up to the total number of controls minus 1 for the last tab position. For example, if a form has three controls that each have a **TabIndex** property, valid **TabIndex** property settings are 0, 1, and 2.

Setting the  **TabIndex** property to an integer less than 0 produces an error.

By default, Microsoft Access assigns a tab order to controls in the order that you create them on a form. Each new control is placed last in the tab order. If you change the setting of a control's  **TabIndex** property to adjust the tab order, Microsoft Access automatically renumbers the **TabIndex** property setting of other controls to reflect insertions and deletions.

In Form view, invisible or disabled controls remain in the tab order but are skipped when you press the TAB key.

Changing the tab order of other controls on the form doesn't affect what happens when you press a control's access key. For example, if you've created an access key for the label of a text box, the focus will move to the text box whenever you press the label's access key — even if you change the  **TabIndex** property setting for the text box.

If you press an access key for a control such as a label that doesn't have a  **TabIndex** property (and thus isn't in the tab order), the focus moves to the next control in the tab order that can receive the focus.


## Example

The following example reverses the tab order of a command button and a text box. Because TextBox1 was created first, it has a  **TabIndex** property setting of 0 and Command1 has a setting of 1.


```vb
Sub Form_Click() 
 Me!Command1.TabIndex = 0 
 Me!TextBox1.TabIndex = 1 
End Sub
```


## See also


#### Concepts


[OptionGroup Object](optiongroup-object-access.md)

