---
title: TextBox.Change Event (Access)
keywords: vbaac10.chm14196
f1_keywords:
- vbaac10.chm14196
ms.prod: access
api_name:
- Access.TextBox.Change
ms.assetid: adde0a6d-d37a-a457-0dea-f2358adbb665
ms.date: 06/08/2017
---


# TextBox.Change Event (Access)

The  **Change** event occurs when the contents of the specified control changes.


## Syntax

 _expression_. **Change**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

Examples of this event include entering a character directly in the text box or combo box or changing the control's  **Text** property setting by using a macro or Visual Basic.


 **Note**  Setting the value of a control by using a macro or Visual Basic doesn't trigger this event for the control. You must type the data directly into the control, or set the control's  **Text** property.

To run a macro or event procedure when this event occurs, set the  **OnChange** property to the name of the macro or to [Event Procedure].

By running a macro or event procedure when a Change event occurs, you can coordinate data display among controls. You can also display data or a formula in one control and the results in another control.

The Change event doesn't occur when a value changes in a calculated control.


 **Note**  A Change event can cause a cascading event. This occurs when a macro or event procedure that runs in response to the control's Change event alters the control's contents — for example, by changing a property setting that determines the control's value, such as the  **Text** property for a text box. To prevent a cascading event:


- If possible, avoid attaching a Change macro or event procedure to a control that alters the control's contents.
    
- Avoid creating two or more controls having Change events that affect each other — for example, two text boxes that update each other.
    
Changing the data in a text box or combo box by using the keyboard causes keyboard events to occur in addition to control events like the Change event. For example, if you move to a new record and type an ANSI character in a text box in the record, the following events occur in this order:

 **KeyDown** → **KeyPress** → **BeforeInsert** → **Change** → **KeyUp**

The  **BeforeUpdate** and **AfterUpdate** events for the text box or combo box control occur after you have entered the new or changed data in the control and moved to another control (or clicked **Save Record** on the **Records** menu), and therefore after all of the Change events for the control.

In combo boxes for which the  **LimitToList** property is set to Yes, the **NotInList** event occurs after you enter a value that isn't in the list and attempt to move to another control or save the record. It occurs after all the Change events for the combo box. In this case, the BeforeUpdate and AfterUpdate events for the combo box don't occur, because Microsoft Access doesn't accept a value that is not in the list.


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

