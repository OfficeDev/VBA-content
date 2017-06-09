---
title: ListBox.IMESentenceMode Property (Access)
keywords: vbaac10.chm11298
f1_keywords:
- vbaac10.chm11298
ms.prod: access
api_name:
- Access.ListBox.IMESentenceMode
ms.assetid: 877e1766-c378-cf7b-b452-bb8f536980f3
ms.date: 06/08/2017
---


# ListBox.IMESentenceMode Property (Access)





## Syntax

 _expression_. **IMESentenceMode**

 _expression_ A variable that represents a **ListBox** object.


## Remarks

The  **IMESentenceMode** property accepts the[AcImeSentenceMode Enumeration (Access)](acimesentencemode-enumeration-access.md) enumeration.



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


[ListBox Object](listbox-object-access.md)

