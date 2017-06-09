---
title: DragBehavior Property
keywords: fm20.chm2001085
f1_keywords:
- fm20.chm2001085
ms.prod: office
api_name:
- Office.DragBehavior
ms.assetid: 8145cbe3-0e13-0715-1c21-b2f4f2ed7b86
ms.date: 06/08/2017
---


# DragBehavior Property



Specifies whether the system enables the drag-and-drop feature for a  **TextBox** or **ComboBox**.
 **Syntax**
 _object_. **DragBehavior** [= _fmDragBehavior_ ]
The  **DragBehavior** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmDragBehavior_|Optional. Specifies whether the drag-and-drop feature is enabled.|
 **Settings**
The settings for  _fmDragBehavior_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmDragBehaviorDisabled_|0|Does not allow a drag-and-drop action (default).|
| _fmDragBehaviorEnabled_|1|Allows a drag-and-drop action.|
 **Remarks**
If the  **DragBehavior** property is enabled, dragging in a text box or combo box starts a drag-and-drop operation on the selected text. If **DragBehavior** is disabled, dragging in a text box or combo box selects text.
The drop-down portion of a  **ComboBox** does not support drag-and-drop processes, nor does it support selection of list items within the text.
 **DragBehavior** has no effect on a **ComboBox** whose **Style** property is set to **fmStyleDropDownList**.

 **Note**  You can combine the effects of the  **EnterFieldBehavior** property and **DragBehavior** to create a large number of text box styles.


