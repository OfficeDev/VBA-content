---
title: TabStop.Leader Property (Publisher)
keywords: vbapb10.chm5636101
f1_keywords:
- vbapb10.chm5636101
ms.prod: publisher
api_name:
- Publisher.TabStop.Leader
ms.assetid: a788bdc8-8ab3-fcd3-931a-a5b83db93991
ms.date: 06/08/2017
---


# TabStop.Leader Property (Publisher)

Sets or returns a  **PbTabLeaderType** constant that represents the leader character for a tab stop. Read/write.


## Syntax

 _expression_. **Leader**

 _expression_A variable that represents a  **TabStop** object.


### Return Value

PbTabLeaderType


## Remarks

The  **Leader** property value can be one of the **[PbTabLeaderType](pbtableadertype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example changes the leader tab character of the selected paragraphs to dashes. This example assumes that the selected paragraph contains at least one tab stop.


```vb
Sub SetLeaderTab() 
 Selection.TextRange.ParagraphFormat _ 
 .Tabs(1).Leader = pbTabLeaderDashes 
End Sub
```

This example changes the leader tab character of the first paragraph in the specified text range to an underline. This example assumes that the specified paragraph contains at least one tab stop.




```vb
Sub SetNewTabLeader() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.Paragraphs(1) _ 
 .ParagraphFormat.Tabs(1).Leader = pbTabLeaderLine 
End Sub
```


