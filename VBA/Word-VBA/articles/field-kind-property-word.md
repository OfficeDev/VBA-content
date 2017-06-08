---
title: Field.Kind Property (Word)
keywords: vbawd10.chm154075139
f1_keywords:
- vbawd10.chm154075139
ms.prod: word
api_name:
- Word.Field.Kind
ms.assetid: 8da8e1a1-5e4c-96fd-7ce3-f650433c1ed1
ms.date: 06/08/2017
---


# Field.Kind Property (Word)

Returns the type of link for a  **Field** object. Read-only **[WdFieldKind](wdfieldkind-enumeration-word.md)** .


## Syntax

 _expression_ . **Kind**

 _expression_ Required. A variable that represents a **[Field](field-object-word.md)** object.


## Example

This example updates all warm link fields in the active document.


```vb
For Each aField In ActiveDocument.Fields 
 If aField.Kind = wdFieldKindWarm Then aField.Update 
Next aField
```


## See also


#### Concepts


[Field Object](field-object-word.md)

