---
title: ComboBox.AfterUpdate Event (Access)
keywords: vbaac10.chm14212
f1_keywords:
- vbaac10.chm14212
ms.prod: access
api_name:
- Access.ComboBox.AfterUpdate
ms.assetid: 89b45f0c-5ab1-889e-bd26-a34281b49b9e
ms.date: 06/08/2017
---


# ComboBox.AfterUpdate Event (Access)

The  **AfterUpdate** event occurs after changed data in a control or record is updated.


## Syntax

 _expression_. **AfterUpdate**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

Changing data in a control by using Visual Basic or a macro containing the SetValue action doesn't trigger these events for the control. However, if you then move to another record or save the record, the form's  **AfterUpdate** event does occur.

To run a macro or event procedure when this event occurs, set the  **AfterUpdate** property to the name of the macro or to [Event Procedure].

The  **AfterUpdate** event is triggered when a control or record is updated. Within a record, changed data in each control is updated when the control loses the focus or when the user presses ENTER or TAB.

When you enter new or changed data in a control on a form and then move to another record or save the record by clicking  **Save Record** on the **Records** menu, the **AfterUpdate** event for the form occur immediately after the **AfterUpdate** event for the control. When you move to a different record, the **Exit** and **LostFocus** events for the control occur, followed by the **Current** event for the record you moved to, and the **Enter** and **GotFocus** events for the first control in this record. To run the **AfterUpdate** macro or event procedure without running the **Exit** and **LostFocus** macros or event procedures, save the record by using the **Save Record** command on the **Records** menu.

 **AfterUpdate** macros and event procedures run only if you change the data in a control. This event does not occur when a value changes in a calculated control. **AfterUpdate** macros and event procedures for a form run only if you change the data in one or more controls in the record.

For bound controls, the  **OldValue** property isn't set to the updated value until after the **AfterUpdate** event for the form occurs. Even if the user enters a new value in the control, the **OldValue** property setting isn't changed until the data is saved (the record is updated). If you cancel an update, the value of the **OldValue** property replaces the existing value in the control.


 **Note**  To perform simple validations, or more complex validations such as requiring a value in a field or validating more than one control on a form, you can use the  **ValidationRule** property for controls and the **ValidationRule** and **Required** properties for fields and records in tables.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Forms: Populate Controls/Text Boxes Based on Combobox Selection](http://www.utteraccess.com/wiki/index.php/Forms:_Populate_Controls/Text_Boxes_Based_on_Combobox_Selection)
    

## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[ComboBox Object](combobox-object-access.md)

