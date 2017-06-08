---
title: ComboBox.CurX Property (Outlook Forms Script)
keywords: olfm10.chm2001040
f1_keywords:
- olfm10.chm2001040
ms.prod: outlook
ms.assetid: ecd78eb7-2ccf-29c3-00c2-641c1f5a4c78
ms.date: 06/08/2017
---


# ComboBox.CurX Property (Outlook Forms Script)

Returns or sets a  **Long** that represents the current horizontal position of the insertion point in a multiline **[ComboBox](combobox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **CurX**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The  **CurX** property applies to a multiline **ComboBox**. The return value is valid when the object has the focus.

You can use  **[CurTargetX](combobox-curtargetx-property-outlook-forms-script.md)** and **CurX** to position the insertion point as the user scrolls through the contents of a multiline **ComboBox**. When the user moves the insertion point to another line of text by scrolling the content of the object,  **CurTargetX** specifies the preferred position for the insertion point. **CurX** is set to this value if the line of text is longer than the value of **CurTargetX**. Otherwise,  **CurX** is set to the end of the line of text.


