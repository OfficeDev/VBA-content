---
title: TabStops.ClearAll Method (Word)
keywords: vbawd10.chm156565605
f1_keywords:
- vbawd10.chm156565605
ms.prod: word
api_name:
- Word.TabStops.ClearAll
ms.assetid: 757bf3e9-5641-8e78-b209-1512087fcf6d
ms.date: 06/08/2017
---


# TabStops.ClearAll Method (Word)

Clears all the custom tab stops from the specified paragraphs.


## Syntax

 _expression_ . **ClearAll**

 _expression_ Required. A variable that represents a **[TabStops](tabstops-object-word.md)** collection.


## Remarks

To clear an individual tab stop, use the  **Clear** method of the **TabStop** object. The **ClearAll** method doesn't clear the default tab stops. To manipulate the default tab stops, use the **DefaultTabStop** property for the document.








## Example

This example clears all the custom tab stops in the active document.


```vb
ActiveDocument.Paragraphs.TabStops.ClearAll
```


## See also


#### Concepts


[TabStops Collection Object](tabstops-object-word.md)

