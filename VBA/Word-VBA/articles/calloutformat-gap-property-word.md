---
title: CalloutFormat.Gap Property (Word)
keywords: vbawd10.chm163905643
f1_keywords:
- vbawd10.chm163905643
ms.prod: word
api_name:
- Word.CalloutFormat.Gap
ms.assetid: 0541a8a6-7eac-d03b-8438-c6d2918237fd
ms.date: 06/08/2017
---


# CalloutFormat.Gap Property (Word)

Returns or sets the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write  **Single** .


## Syntax

 _expression_ . **Gap**

 _expression_ A variable that represents a **[CalloutFormat](calloutformat-object-word.md)** object.


## Example

This example sets the distance between the callout line and the text bounding box to 3 points for the first shape on the active document. For the example to work, the first shape must be a callout.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
docActive.Shapes(1).Callout.Gap = 3
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-word.md)

