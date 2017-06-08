---
title: Tags Object (Publisher)
keywords: vbapb10.chm4718591
f1_keywords:
- vbapb10.chm4718591
ms.prod: publisher
api_name:
- Publisher.Tags
ms.assetid: 76cccc1e-4623-af8b-f0f8-e6cc245b94fd
ms.date: 06/08/2017
---


# Tags Object (Publisher)

A collection of  **Tag** objects that represents tags or custom properties applied to a shape, shape range, page, or publication.
 


## Example

Use the  **[Tags](shape-tags-property-publisher.md)** property to access the **Tags** collection. Use the **[Add](tags-add-method-publisher.md)** method of the **Tags** collection to add a **Tag** object to a shape, shape range, page, or publication. This example adds a tag to each oval shape on the first page of the active publication.
 

 

```
Sub AddNewTag() 
 Dim shp As Shape 
 With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If InStr(1, shp.Name, "Oval") > 0 Then 
 shp.Tags.Add Name:="Shape", Value:="Oval" 
 End If 
 Next shp 
 End With 
End Sub
```

Use the  **[Count](tags-count-property-publisher.md)** property to determine if a shape, shape range, page, or publication contains one or more **Tag** objects. This example fills all shapes on the first page of the active publication if the shape's first tag has a value of Oval.
 

 



```
Sub FormatTaggedShapes() 
 Dim shp As Shape 
 With ActiveDocument.Pages(1) 
 For Each shp In .Shapes 
 If shp.Tags.Count > 0 Then 
 If shp.Tags(1).Value = "Oval" Then 
 shp.Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 End If 
 End If 
 Next shp 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](tags-add-method-publisher.md)|
|[Item](tags-item-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](tags-application-property-publisher.md)|
|[Count](tags-count-property-publisher.md)|
|[Parent](tags-parent-property-publisher.md)|

