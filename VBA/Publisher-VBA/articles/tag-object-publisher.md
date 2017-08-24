---
title: Tag Object (Publisher)
keywords: vbapb10.chm4784127
f1_keywords:
- vbapb10.chm4784127
ms.prod: publisher
api_name:
- Publisher.Tag
ms.assetid: f485d2cc-8e39-5aa3-d407-8c14401ec8bd
ms.date: 06/08/2017
---


# Tag Object (Publisher)

Represents a tag or a custom property that you can create for a shape, shape range, page, or publication. Each  **Tag** object contains the name of a custom property and a value for that property. **Tag** objects are members of the **[Tags](tags-object-publisher.md)** collection. Create a tag when you want to be able to selectively work with specific members of a collection, based on an attribute that isn't already represented by a built-in property.
 


## Example

Use the  **[Item](tags-item-method-publisher.md)** method of the **[Tags](tags-object-publisher.md)** collection to return a **Tag** object. This example fills all shapes on the first page of the active publication if the shape's first tag has a value of Oval.
 

 

```
Sub FormatTaggedShapes() 
 Dim shp As Shape 
 With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If shp.Tags.Count > 0 Then 
 If shp.Tags.Item(1).Value = "Oval" Then 
 shp.Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 End If 
 End If 
 Next 
 End With 
End Sub
```

Use the  **[Add](tags-add-method-publisher.md)** method to add a Tag object. This example adds a tag to all oval shapes in the active publication.
 

 



```
Sub TagShapes() 
 Dim shp As Shape 
 With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If InStr(1, shp.Name, "Oval") > 0 Then 
 shp.Tags.Add Name:="Oval", Value:="This is an oval shape." 
 End If 
 Next shp 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](tag-delete-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](tag-application-property-publisher.md)|
|[Name](tag-name-property-publisher.md)|
|[Parent](tag-parent-property-publisher.md)|
|[Value](tag-value-property-publisher.md)|

