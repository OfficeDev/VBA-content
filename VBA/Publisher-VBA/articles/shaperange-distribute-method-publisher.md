---
title: ShapeRange.Distribute Method (Publisher)
keywords: vbapb10.chm2294017
f1_keywords:
- vbapb10.chm2294017
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Distribute
ms.assetid: a145fb46-d7b6-bc3c-b7fd-cdb892fda179
ms.date: 06/08/2017
---


# ShapeRange.Distribute Method (Publisher)

Evenly distributes the shapes in the specified shape range.


## Syntax

 _expression_. **Distribute**( **_DistributeCmd_**,  **_RelativeTo_**)

 _expression_A variable that represents a  **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|DistributeCmd|Required| **MsoDistributeCmd**|Specifies whether shapes are to be distributed horizontally or vertically.|
|RelativeTo|Required| **MsoTriState**|Specifies whether to distribute the shapes evenly over the entire horizontal or vertical space on the page or within the horizontal or vertical space that the range of shapes originally occupies.|

## Remarks

Shapes are distributed so that there is an equal amount of space between one shape and the next. If the shapes are so large that they overlap when distributed over the available space, they are distributed so that there is an equal amount of overlap between one shape and the next.

The DistributeCmd parameter can be one of the following  **MsoDistributeCmd** constants declared in the Microsoft Office type library.



| **msoDistributeHorizontally**|
| **msoDistributeVertically**|
The RelativeTo parameter can be one of the  **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| Distribute the shapes within the horizontal or vertical space that the range of shapes originally occupies.|
| **msoTrue**|Distribute the shapes evenly over the entire horizontal or vertical space on the page.|
When RelativeTo is  **msoTrue**, shapes are distributed so that the distance between the two outer shapes and the edges of the page is the same as the distance between one shape and the next. If the shapes must overlap, the two outer shapes are moved to the edges of the page.

When RelativeTo is  **msoFalse**, the two outer shapes are not moved; only the positions of the inner shapes are adjusted.

The z-order of shapes is unaffected by this method.


## Example

This example defines a shape range that contains all the AutoShapes on the first page of the active publication and then horizontally distributes the shapes in this range.


```vb
' Number of shapes on the page. 
Dim intShapes As Integer 
' Number of AutoShapes on the page. 
Dim intAutoShapes As Integer 
' An array of the names of the AutoShapes. 
Dim arrAutoShapes() As String 
' A looping variable. 
Dim shpLoop As Shape 
' A placeholder variable for the range containing AutoShapes. 
Dim shpRange As ShapeRange 
 
With ActiveDocument.Pages(1).Shapes 
 ' Count all the shapes on the page. 
 intShapes = .Count 
 
 ' Proceed only if there's at least one shape. 
 If intShapes > 1 Then 
 intAutoShapes = 0 
 ReDim arrAutoShapes(1 To intShapes) 
 
 ' Loop through the shapes on the page and add the names 
 ' of any AutoShapes to an array. 
 For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.Type = msoAutoShape Then 
 intAutoShapes = intAutoShapes + 1 
 arrAutoShapes(intAutoShapes) = shpLoop.Name 
 End If 
 Next shpLoop 
 
 ' Proceed only if there's at least one AutoShape. 
 If intAutoShapes > 1 Then 
 ReDim Preserve arrAutoShapes(1 To intAutoShapes) 
 
 ' Create a shape range containing all the AutoShapes. 
 Set shpRange = .Range(Index:=arrAutoShapes) 
 
 ' Distribute the AutoShapes horizontally 
 ' in the space they already occupy. 
 shpRange.Distribute _ 
 DistributeCmd:=msoDistributeHorizontally, RelativeTo:=msoFalse 
 End If 
 End If 
End With 

```


