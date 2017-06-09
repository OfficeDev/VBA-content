---
title: CurTargetX Property
keywords: fm20.chm5225027
f1_keywords:
- fm20.chm5225027
ms.prod: office
api_name:
- Office.CurTargetX
ms.assetid: b0365f58-22db-34d2-9751-6c9d36598e08
ms.date: 06/08/2017
---


# CurTargetX Property



Retrieves the preferred horizontal position of the insertion point in a multiline  **TextBox** or **ComboBox**.
 **Syntax**
 _object_. **CurTargetX**
The  **CurTargetX** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Return Values**
The  **CurTargetX** property retrieves the preferred position, measured in himetric units. A himetric is 0.0001 meter.
 **Remarks**
The [target](glossary-vba.md) position is relative to the left edge of the control. If the length of a line is less than the value of the **CurTargetX** property, you can place the insertion point at the end of the the line. The value of **CurTargetX** changes when the user sets the insertion point or when the **CurX** property is set. **CurTargetX** is read-only.
The return value is valid when the object has [focus](vbe-glossary.md).
You can use  **CurTargetX** and **CurX** to move the insertion point as the user scrolls through the contents of a multiline **TextBox** or **ComboBox**. When the user moves the insertion point to another line of text by scrolling the content of the object, **CurTargetX** specifies the preferred position for the insertion point. **CurX** is set to this value if the line of text is longer than the value of **CurTargetX**. Otherwise, **CurX** is set to the end of the line of text.

