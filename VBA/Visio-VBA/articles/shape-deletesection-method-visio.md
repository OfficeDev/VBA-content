---
title: Shape.DeleteSection Method (Visio)
keywords: vis_sdr.chm11216175
f1_keywords:
- vis_sdr.chm11216175
ms.prod: visio
api_name:
- Visio.Shape.DeleteSection
ms.assetid: e07981f3-5efe-f4ad-0517-1af4913c3f70
ms.date: 06/08/2017
---


# Shape.DeleteSection Method (Visio)

Deletes a ShapeSheet section.


## Syntax

 _expression_ . **DeleteSection**( **_Section_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The index of the section to delete.|

### Return Value

Nothing


## Remarks

When you delete a ShapeSheet section, all rows in the section are automatically deleted. If the specified section does not exist, nothing is deleted and no error is generated.

If a Geometry section is deleted, any subsequent Geometry sections shift up because they are indexed and no gaps can exist in an indexed range.

You can delete any section except the section represented by  **visSectionObject** (although you can delete rows within that section).

Section index values are declared in the Visio type library in  **[VisSectionIndices](vissectionindices-enumeration-visio.md)** .


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to delete a ShapeSheet section.


```vb
Public Sub DeleteSection_Example() 
 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 
 'Get the active page. 
 Set vsoPage = ActivePage 
 
 'If there isn't an active page, set the Page object 
 'to the first page of the active document. 
 If vsoPage Is Nothing Then 
 Set vsoPage = ActiveDocument.Pages(1) 
 End If 
 
 'Draw a rectangle on the active page. 
 Set vsoShape = vsoPage.DrawRectangle(1, 5, 5, 1) 
 
 'Add a scratch section to the ShapeSheet of the rectangle. 
 vsoShape.AddSection visSectionScratch 
 
 'Delete the scratch section from the ShapeSheet. 
 vsoShape.DeleteSection visSectionScratch 
 
End Sub
```


