---
title: Shape.MoveIntoTextFlow Method (Publisher)
keywords: vbapb10.chm2228356
f1_keywords:
- vbapb10.chm2228356
ms.prod: publisher
api_name:
- Publisher.Shape.MoveIntoTextFlow
ms.assetid: d8a2af57-f974-717e-0d97-c8a3aee16f01
ms.date: 06/08/2017
---


# Shape.MoveIntoTextFlow Method (Publisher)

Moves a given shape into the text flow defined by  ** [TextRange Object](textrange-object-publisher.md)**. The shape will always be inserted inline at the beginning of the text flow.


## Syntax

 _expression_. **MoveIntoTextFlow**( **_Range_**)

 _expression_A variable that represents a  **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Range|Required| **TextRange**|The range of text before which the given shape is inserted.|

### Return Value

Nothing


## Remarks

The  **MoveIntoTextFlow** method will fail if the shape to be moved is already inline or if it is not a valid inline shape type. Invalid inline shape types include:


- Inline shapes
    
- Grouped shapes
    
- HTML fragments
    
- Smart objects
    
- Chained text boxes
    



## Example

The following example checks if the second shape on the second page of the publication is inline, and if it is not, inserts it inline at the beginning of the text flow of the given text range. 


```vb
Dim theShape As Shape 
Dim theRange As TextRange 
 
Set theRange = ActiveDocument.Pages(2).Shapes(1).TextFrame.TextRange 
Set theShape = ActiveDocument.Pages(2).Shapes(2) 
 
If Not theShape.IsInline = msoTrue Then 
 theShape.MoveIntoTextFlow Range:=theRange 
End If 

```


