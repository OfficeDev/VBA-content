---
title: TextBox.IMESentenceMode Property (Access)
keywords: vbaac10.chm11137
f1_keywords:
- vbaac10.chm11137
ms.prod: access
api_name:
- Access.TextBox.IMESentenceMode
ms.assetid: 399a28d4-83a9-33d2-5f00-4f388efe048b
ms.date: 06/08/2017
---


# TextBox.IMESentenceMode Property (Access)





## Syntax

 _expression_. **IMESentenceMode**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

The  **IMESentenceMode** property uses the following settings.



|**Setting**|**Description**|**Visual Basic**|
|:-----|:-----|:-----|
|Normal|(Default) Set IME Sentence Mode to ?Normal? mode.|0|
|Plural|Set IME Sentence Mode to ?Plural? mode.|1|
|Speaking|Set IME Sentence Mode to ?Speaking? mode.|2|
|No Conversion|Doesn?t set IME Sentence Mode.|3|
 **Normal mode**

Use this mode when creating a literary Japanese document.

 **Plural mode**

Use this mode when entering name or address data. In this mode, two additional dictionaries are available. The ?Biographical/Geographical Dictionary? contains names not covered in the normal dictionary and the ?Postal Code Dictionary?, useful in creating addresses. (Factory setting.)

 **Speaking mode**

Use this mode when entering data that contains conversational language.

 **No Conversion**

In this mode, inputted characters are settled without conversion.


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

