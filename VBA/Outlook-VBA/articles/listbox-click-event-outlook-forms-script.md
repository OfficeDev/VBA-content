---
title: ListBox.Click Event (Outlook Forms Script)
ms.prod: outlook
ms.assetid: a3b32670-d20c-a5cc-d236-041cbe155779
ms.date: 06/08/2017
---


# ListBox.Click Event (Outlook Forms Script)

Occurs when the user definitively selects a value for the control that has more than one possible value.


## Syntax

 _expression_. **Click**

 _expression_A variable that represents a  **ListBox** object.


## Remarks

The following are examples of actions that initiate the  **Click** event of the specified control:


- Clicking a blank area of a form or a disabled control (other than a list box) on the form.
    
- Clicking a control with the left mouse button (left-clicking).
    
- Pressing a control's accelerator key.
    


For some controls, the  **Click** event occurs when the **Value** property changes. However, using the **PropertyChange** or **CustomPropertyChange** event is the preferred technique for detecting a new value for a property. The following are examples of actions that initiate the **Click** event due to assigning a new value to a control: selecting a value for a **[ListBox](listbox-object-outlook-forms-script.md)** so that it unquestionably matches an item in the control's drop-down list. For example, if a list is not sorted, the first match for characters typed in the edit region may not be the only match in the list, so choosing such a value does not initiate the **Click** event. In a sorted list, you can use entry-matching to ensure that a selected value is a unique match for text the user types.

The  **Click** event is not initiated when **Value** is set to **Null**.

Left-clicking changes the value of a control, thus it initiates the  **Click** event. Right-clicking does not change the value of the control, so it does not initiate the **Click** event.

If you bind a  **ListBox** to a field, then the **Click** event does not fire. You need to use the **PropertyChange** or **CustomPropertyChange** event to detect the change via code, as in the following code sample:




```vb
Sub Item_PropertyChange(ByVal Name) 
Set MyListBox = Item.GetInspector.ModifiedFormPages("Message").Controls("ListBox1") 
Select Case Name 
 Case "Mileage" 
 Item.CC = MyListBox.Value 
 Item.Subject = MyListBox.Value 
 Case Else 
End Select 
End Sub
```


