---
title: Global.NormalTemplate Property (Word)
keywords: vbawd10.chm163119112
f1_keywords:
- vbawd10.chm163119112
ms.prod: word
api_name:
- Word.Global.NormalTemplate
ms.assetid: ddfcd859-5d4c-e5f7-a04e-70102c1780d2
ms.date: 06/08/2017
---


# Global.NormalTemplate Property (Word)

Returns a  **Template** object that represents the Normal template.


## Syntax

 _expression_ . **NormalTemplate**

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


## Example

This example inserts the AutoText entry named "Test" from the Normal template, if this entry is contained in the  **AutoTextEntries** collection.


```vb
For Each entry In NormalTemplate.AutoTextEntries 
 If entry.Name = "Test" Then entry.Insert Where:=Selection.Range 
Next entry
```

This example saves the Normal template if changes have been made to it.




```vb
If NormalTemplate.Saved = False Then NormalTemplate.Save
```


## See also


#### Concepts


[Global Object](global-object-word.md)

