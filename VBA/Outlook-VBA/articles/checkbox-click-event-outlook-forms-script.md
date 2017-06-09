---
title: CheckBox.Click Event (Outlook Forms Script)
keywords: olfm10.chm2000070
f1_keywords:
- olfm10.chm2000070
ms.prod: outlook
ms.assetid: 186f0164-0d7d-0068-b8ec-2e1bc6e561cd
ms.date: 06/08/2017
---


# CheckBox.Click Event (Outlook Forms Script)

Occurs when the user clicks inside the control.


## Syntax

 _expression_. **Click**

 _expression_A variable that represents a  **CheckBox** object.


## Remarks

The following are examples of actions that initiate the  **Click** event of the specified control:


- Clicking a blank area of a form or a disabled control (other than a list box) on the form.
    
- Clicking a control with the left mouse button (left-clicking).
    
- Pressing a control's accelerator key.
    


For some controls, the  **Click** event occurs when the **Value** property changes. However, using the **PropertyChange** or **CustomPropertyChange** event is the preferred technique for detecting a new value for a property. The following are examples of actions that initiate the **Click** event due to assigning a new value to a control: clicking a **[CheckBox](checkbox-object-outlook-forms-script.md)**, pressing the  **SPACEBAR** when the check box has the focus, pressing the accelerator key, or changing the value of the control in code.

The  **Click** event is not initiated when **Value** is set to **Null**.

Left-clicking changes the value of a control, thus it initiates the  **Click** event. Right-clicking does not change the value of the control, so it does not initiate the **Click** event.

If you bind a  **CheckBox** to a field, then the **Click** event does not fire. You need to use the **PropertyChange** or **CustomPropertyChange** event to detect the change via code, as in the following code sample:




```vb
Sub Item_PropertyChange(ByVal Name) 
Set MyListBox = Item.GetInspector.ModifiedFormPages("Message").Controls("CheckBox1") 
Select Case Name 
    Case "Mileage" 
        Item.CC = MyCheckBox.Value 
        Item.Subject = MyCheckBox.Value 
    Case Else 
End Select 
End Sub
```


