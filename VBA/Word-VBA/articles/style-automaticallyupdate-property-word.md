---
title: Style.AutomaticallyUpdate Property (Word)
keywords: vbawd10.chm153878541
f1_keywords:
- vbawd10.chm153878541
ms.prod: word
api_name:
- Word.Style.AutomaticallyUpdate
ms.assetid: 6b224938-9519-5cb3-4ca5-ca9e465432e9
ms.date: 06/08/2017
---


# Style.AutomaticallyUpdate Property (Word)

 **True** if the style is automatically redefined based on the selection. Read/write **Boolean** .


## Syntax

 _expression_ . **AutomaticallyUpdate**

 _expression_ A variable that represents a **[Style](style-object-word.md)** object.


## Remarks

If the  **AutomaticallyUpdate** property is set to **False** , Microsoft Word prompts for confirmation before redefining the style based on the selection. A style can be redefined when it is applied to a selection that has the same style but different manual formatting. The AutomaticallyUpdate property applies to paragraph styles only.


## Example

This example creates a style named "Style1" that can be redefined without the need for confirmation.


```vb
Dim docNew as Document 
Dim styleNew as Style 
 
Set docNew = Documents.Add 
Set styleNew = docNew.Styles.Add("Style1") 
 
With styleNew 
 .BaseStyle = docNew.Styles(wdStyleNormal) 
 .ParagraphFormat.LineSpacingRule = wdLineSpaceDouble 
 .AutomaticallyUpdate = True 
End With
```


## See also


#### Concepts


[Style Object](style-object-word.md)

