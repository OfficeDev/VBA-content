---
title: TabStops.ClearAll Method (Publisher)
keywords: vbapb10.chm5570564
f1_keywords:
- vbapb10.chm5570564
ms.prod: publisher
api_name:
- Publisher.TabStops.ClearAll
ms.assetid: bb7e2a0e-c044-872d-aa74-2683886e77a6
ms.date: 06/08/2017
---


# TabStops.ClearAll Method (Publisher)

Clears all the custom tab stops from the specified paragraphs.


## Syntax

 _expression_. **ClearAll**

 _expression_A variable that represents a  **TabStops** object.


## Remarks

To clear an individual tab stop, use the  **[Clear](tabstop-clear-method-publisher.md)** method of the  **[TabStop](tabstop-object-publisher.md)** object. The  **ClearAll** method doesn't clear the default tab stops. To manipulate the default tab stops, use the **[DefaultTabStop](document-defaulttabstop-property-publisher.md)** property for the document.


## Example

This example clears all the custom tab stops in the first shape on the first page of the active publication. This example assumes that the specified shape is a text frame and not another type of shape.


```vb
Sub ClearAllTabStops() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat.Tabs.ClearAll 
End Sub
```


