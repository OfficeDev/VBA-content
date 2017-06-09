---
title: CoAuthoring.Conflicts Property (Word)
keywords: vbawd10.chm254869511
f1_keywords:
- vbawd10.chm254869511
ms.prod: word
api_name:
- Word.CoAuthoring.Conflicts
ms.assetid: bd6aab5d-5342-ee1b-c5af-1f67753d55fc
ms.date: 06/08/2017
---


# CoAuthoring.Conflicts Property (Word)

Returns a  **[Conflicts](conflicts-object-word.md)** collection that represents all the conflicts in a document. Read-only.


## Syntax

 _expression_ . **Conflicts**

 _expression_ An expression that returns a **[CoAuthoring](coauthoring-object-word.md)** object.


## Example

The following code example gets the type of each conflict in the active document. The  **[Type](conflict-type-property-word.md)** property uses the **[WdRevisionType](wdrevisiontype-enumeration-word.md)** enumeration to specify the conflict type.


```vb
Dim conf As Conflict 
 
For Each conf In ActiveDocument.CoAuthoring.Conflicts 
    MsgBox conf.Type 
Next conf 

```


## See also


#### Concepts


[CoAuthoring Object](coauthoring-object-word.md)

