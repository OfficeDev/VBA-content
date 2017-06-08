---
title: Shape.BorderArt Property (Publisher)
keywords: vbapb10.chm5308675
f1_keywords:
- vbapb10.chm5308675
ms.prod: publisher
api_name:
- Publisher.Shape.BorderArt
ms.assetid: dcc0ceb4-ef69-ffd3-e510-13dcb8d06832
ms.date: 06/08/2017
---


# Shape.BorderArt Property (Publisher)

Returns a  **[BorderArtFormat](borderartformat-object-publisher.md)** object that represents the BorderArt type applied to the specified shape. Returns "Permission Denied" if BorderArt has not been applied to the shape. Read-only.


## Syntax

 _expression_. **BorderArt**

 _expression_A variable that represents a  **Shape** object.


### Return Value

BorderArtFormat


## Remarks

BorderArt are picture borders that can be applied to text boxes, picture frames, or rectangles. 

Use the  **BorderArt** property to apply, change, and remove BorderArt from shapes in publications.


## Example

The following example tests for the existence of BorderArt on each shape for each page of the active publication. If BorderArt exists, it is deleted.


```vb
Sub DeleteBorderArt() 
 
Dim anyPage As Page 
Dim anyShape As Shape 
 
For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .Delete 
 End If 
 End With 
 Next anyShape 
 Next anyPage 
End Sub
```


