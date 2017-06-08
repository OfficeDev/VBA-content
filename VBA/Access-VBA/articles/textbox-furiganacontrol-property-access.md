---
title: TextBox.FuriganaControl Property (Access)
keywords: vbaac10.chm11049
f1_keywords:
- vbaac10.chm11049
ms.prod: access
api_name:
- Access.TextBox.FuriganaControl
ms.assetid: 7d70cffa-06bb-fa9d-686a-0031558aa5a3
ms.date: 06/08/2017
---


# TextBox.FuriganaControl Property (Access)





## Syntax

 _expression_. **FuriganaControl**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

The setting value is the name of the control for entering furigana.

If the  **FuriganaControl** property is set in the control, furigana will automatically be created, and can be displayed in a separately designated control. Only if a control name in the same form is set with the **FuriganaControl** property will the form run properly when executed. If text is entered in a control other than the designated control name in the same form, an error will occur. The type of furigana characters is determined by the **IMEMode/KanjiConversionMode** property settings in the control.

 **FuriganaControl property in ADP**

When you use  **FuriganaControl** property in ADP file, be sure to change the data type from CHAR/NCHAR to VARCHAR/NVARCHR. Otherwise, you cannot insert any furigana string into the target field. The **FuriganaControl** property inserts furigana strings to an existing target field, but if the data type definition of the field stays as CHAR/NCHAR, any string insertion fails because the field length is fixed, which result in an error message.


 **Note**  If you enter text in the target control, the furigana of the newly entered text is automatically added to the end of the designated target control content. Even if the text of the target control is revised or deleted, the characters before the change in the target control will not be revised or deleted. Changing the content of the target control revises the text in the furigana control as necessary. The  **FuriganaControl** property will not run if text is pasted into the target control.


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

