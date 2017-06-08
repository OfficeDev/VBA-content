---
title: OptionButton.Click Event (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 96bb2ed3-ded1-86e2-f39d-2d651f160ce4
ms.date: 06/08/2017
---


# OptionButton.Click Event (Outlook Forms Script)

Occurs when the user definitively selects a value for the control that has more than one possible value, or when the value changes to  **True**.


## Syntax

 _expression_. **Click**

 _expression_A variable that represents an  **OptionButton** object.


## Remarks

The following are examples of actions that initiate the  **Click** event of the specified control:


- Clicking a blank area of a form or a disabled control (other than a list box) on the form.
    
- Clicking a control with the left mouse button (left-clicking).
    
- Pressing a control's accelerator key.
    


For some controls, the  **Click** event occurs when the **Value** property changes. However, using the **PropertyChange** or **CustomPropertyChange** event is the preferred technique for detecting a new value for a property. The following are examples of actions that initiate the **Click** event due to assigning a new value to a control: changing the value of an **[OptionButton](optionbutton-object-outlook-forms-script.md)** to **True**, and setting one  **OptionButton** in a group to **True** sets all other buttons in the group to **False**, but the  **Click** event occurs only for the button whose value changes to **True**.

The  **Click** event is not initiated when **Value** is set to **Null**.

Left-clicking changes the value of a control, thus it initiates the  **Click** event. Right-clicking does not change the value of the control, so it does not initiate the **Click** event.

If you bind an  **OptionButton** to a field, then the **Click** event does not fire. You need to use the **PropertyChange** or **CustomPropertyChange** event to detect the change via code, as in the following code sample:




```vb
Sub Item_PropertyChange(ByVal Name) 
Set MyListBox = Item.GetInspector.ModifiedFormPages("Message").Controls("OptionButton1") 
Select Case Name 
    Case "Mileage" 
        Item.CC = MyOptionButton.Value 
        Item.Subject = MyOptionButton.Value 
    Case Else 
End Select 
End Sub
```


