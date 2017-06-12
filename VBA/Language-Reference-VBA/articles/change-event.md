---
title: Change Event
keywords: fm20.chm5224938
f1_keywords:
- fm20.chm5224938
ms.prod: office
api_name:
- Office.Change
ms.assetid: 4bf23772-5ae0-dc1d-1152-b7ea01f7e702
ms.date: 06/08/2017
---


# Change Event



Occurs when the  **Value** property changes.
 **Syntax**
 **Private Sub**_object_ _**Change( )**
The  **Change** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Settings**
The Change event occurs when the setting of the  **Value** property changes, regardless of whether the change results from execution of code or a user action in the interface.
Here are some examples of actions that change the  **Value** property:


- Clicking a  **CheckBox**, **OptionButton**, or **ToggleButton**.
    
- Entering or selecting a new text value for a  **ComboBox**, **ListBox**, or **TextBox**.
    
- Selecting a different tab on a  **TabStrip**.
    
- Moving the scroll box in a  **ScrollBar**.
    
- Clicking the up arrow or down arrow on a  **SpinButton**.
    
- Selecting a different page on a  **MultiPage**.
    

 **Remarks**
The Change event procedure can synchronize or coordinate data displayed among controls. For example, you can use the Change event procedure of a  **ScrollBar** to update the contents of a **TextBox** that displays the value of the **ScrollBar**. Or you can use a Change event procedure to display data and formulas in a work area and results in another area.

 **Note**  In some cases, the Click event may also occur when the  **Value** property changes. However, using the Change event is the preferred technique for detecting a new value for a property.


