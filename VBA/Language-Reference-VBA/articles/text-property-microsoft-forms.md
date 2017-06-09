---
title: Text Property (Microsoft Forms)
keywords: fm20.chm2002070
f1_keywords:
- fm20.chm2002070
ms.prod: office
ms.assetid: 493a251a-3e7b-3a4b-a800-4e9b94d19b4f
ms.date: 06/08/2017
---


# Text Property (Microsoft Forms)



Returns or sets the text in a  **TextBox**. Changes the selected row in a **ComboBox** or **ListBox**.
 **Syntax**
 _object_. **Text** [= _String_ ]
The  **Text** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. A string expression specifying text. The default value is a zero-length string ("").|
 **Remarks**
For a  **TextBox**, any value you assign to the **Text** property is also assigned to the **Value** property.
For a  **ComboBox**, you can use **Text** to update the value of the control. If the value of **Text** matches an existing list entry, the value of the **ListIndex** property (the index of the current row) is set to the row that matches **Text**. If the value of **Text** does not match a row, **ListIndex** is set to -1.
For a  **ListBox**, the value of **Text** must match an existing list entry. Specifying a value that does not match an existing list entry causes an error.
You cannot use  **Text** to change the value of an entry in a **ComboBox** or **ListBox**; use the **Column** or **List** property for this purpose.
The  **ForeColor** property determines the color of the text.

