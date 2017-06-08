---
title: ComboBox.Text Property (Access)
keywords: vbaac10.chm11436
f1_keywords:
- vbaac10.chm11436
ms.prod: access
api_name:
- Access.ComboBox.Text
ms.assetid: 27f99e99-ce53-f5b9-61ed-1ffc4ba9cc4d
ms.date: 06/08/2017
---


# ComboBox.Text Property (Access)

You can use the  **Text** property to set or return the text contained in the text box portion of a combo box. Read/write **String**.


## Syntax

 _expression_. **Text**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

You can set the  **Text** property to the text you want to display in the control. You can also use the **Text** property to read the text currently in the control.


 **Note**  To set or return a control's  **Text** property, the control must have the focus, or an error occurs. To move the focus to a control, you can use the **SetFocus** method or GoToControl action.

While the control has the focus, the  **Text** property contains the text data currently in the control; the **Value** property contains the last saved data for the control. When you move the focus to another control, the control's data is updated, and the **Value** property is set to this new value. The **Text** property setting is then unavailable until the control gets the focus again. If you use the **Save Record** command on the **Records** menu to save the data in the control without moving the focus, the **Text** property and **Value** property settings will be the same.


## Example

The following example uses the  **Text** property to enable a Next button named `btnNext` whenever the user enters text into a text box named `txtName`. Anytime the text box is empty, the Next button is disabled.


```vb
Sub txtName_Change() 
 btnNext.Enabled = Len(Me!txtName.Text &; "")<>0 
End Sub
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

