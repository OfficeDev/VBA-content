---
title: XMLNode.FirstChild Property (Word)
keywords: vbawd10.chm37748745
f1_keywords:
- vbawd10.chm37748745
ms.prod: word
api_name:
- Word.XMLNode.FirstChild
ms.assetid: ce0d09ac-383c-b9b4-3065-c26410b442d5
ms.date: 06/08/2017
---


# XMLNode.FirstChild Property (Word)

Returns a  **DiagramNode** object that represents the first child node of a parent node. Read-only.


## Syntax

 _expression_ . **FirstChild**

 _expression_ Required. A variable that represents a **[XMLNode](xmlnode-object-word.md)** object.


## Remarks

Use the  **LastChild** property to access the last child node. Use the **Root** property to access the parent node in a diagram.




## Example

This example adds an organization chart diagram to the current document, adds three nodes, and assigns the first and last child nodes to variables.


```vb
Sub FirstChild() 
 Dim shpDiagram As Shape 
 Dim dgnRoot As DiagramNode 
 Dim dgnFirstChild As DiagramNode 
 Dim dgnLastChild As DiagramNode 
 Dim intCount As Integer 
 
 'Add organizational chart diagram to the current document 
 Set shpDiagram = ActiveDocument.Shapes.AddDiagram _ 
 (Type:=msoDiagramOrgChart, Left:=10, _ 
 Top:=15, Width:=400, Height:=475) 
 
 'Add the first node to the diagram 
 Set dgnRoot = shpDiagram.DiagramNode.Children.AddNode 
 
 'Add three child nodes 
 For intCount = 1 To 3 
 dgnRoot.Children.AddNode 
 Next intCount 
 
 'Assign the first and last child nodes to variables 
 Set dgnFirstChild = dgnRoot.Children.FirstChild 
 Set dgnLastChild = dgnRoot.Children.LastChild 
End Sub
```


## See also


#### Concepts


[XMLNode Object](xmlnode-object-word.md)

