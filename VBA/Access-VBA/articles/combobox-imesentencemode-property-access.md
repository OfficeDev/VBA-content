---
title: ComboBox.IMESentenceMode Property (Access)
keywords: vbaac10.chm11470
f1_keywords:
- vbaac10.chm11470
ms.prod: access
api_name:
- Access.ComboBox.IMESentenceMode
ms.assetid: f56b97cb-73c9-f5ff-a467-6e7dcd64e613
ms.date: 06/08/2017
---


# ComboBox.IMESentenceMode Property (Access)





## Syntax

 _expression_. **IMESentenceMode**

 _expression_ A variable that represents a **ComboBox** object.


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


[ComboBox Object](combobox-object-access.md)

