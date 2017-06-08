---
title: Document.FindShapesByTag Method (Publisher)
keywords: vbapb10.chm196689
f1_keywords:
- vbapb10.chm196689
ms.prod: publisher
api_name:
- Publisher.Document.FindShapesByTag
ms.assetid: 405a0f39-5892-23da-904a-5188a4340b00
ms.date: 06/08/2017
---


# Document.FindShapesByTag Method (Publisher)

Returns a  **[ShapeRange](shaperange-object-publisher.md)** object that represents the shapes with the specified tag.


## Syntax

 _expression_. **FindShapesByTag**( **_TagName_**)

 _expression_A variable that represents a  **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|TagName|Required| **String**|The name of the tag.|

### Return Value

ShapeRange


## Example

This example adds two shapes to the first page of the active publication, assigns each a tag, and then enters the name of each tag into the text frame of its assigned shape.


```vb
Sub FindShape() 
 Dim strTag1 As String 
 Dim strTag2 As String 
 
 With ActiveDocument.Pages(1).Shapes 
 With .AddShape(Type:=msoShape5pointStar, Left:=50, _ 
 Top:=50, Width:=75, Height:=75) 
 strTag1 = .Tags.Add(Name:="Star", _ 
 Value:="This is a star.").Name 
 End With 
 
 With .AddShape(Type:=msoShapeHeart, Left:=100, _ 
 Top:=100, Width:=75, Height:=75) 
 strTag2 = .Tags.Add(Name:="Heart", _ 
 Value:="This is a heart.").Name 
 End With 
 End With 
 
 With ActiveDocument 
 .FindShapesByTag(TagName:=strTag1).TextFrame _ 
 .TextRange.Text = strTag1 
 .FindShapesByTag(TagName:=strTag2).TextFrame _ 
 .TextRange.Text = strTag2 
 End With 
End Sub
```


