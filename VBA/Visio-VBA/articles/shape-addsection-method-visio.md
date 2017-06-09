---
title: Shape.AddSection Method (Visio)
keywords: vis_sdr.chm11216060
f1_keywords:
- vis_sdr.chm11216060
ms.prod: visio
api_name:
- Visio.Shape.AddSection
ms.assetid: 64396db4-8361-ece9-b029-24d62ba0a290
ms.date: 06/08/2017
---


# Shape.AddSection Method (Visio)

Adds a new section to a ShapeSheet spreadsheet.


## Syntax

 _expression_ . **AddSection**( **_Section_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The type of section to add.|

### Return Value

Integer


## Remarks

The  **AddSection** method is frequently used to add one or more Geometry sections to a shape. You can also use **AddSection** to add other sections to a shape such as Scratch, Controls, Connection Points, Actions, User-Defined Cells and ShapeData. The **AddSection** method returns the logical index of the added section.

The sections that you can add to a shape correspond to the choices shown by the  **Insert Section** dialog box when the shape is displayed in a ShapeSheet window.

If you try to add a non-Geometry section to a shape that already has that section, the  **AddSection** method raises an exception. Use the **SectionExists** property to determine if a shape already has a section with a given logical index.

A new section has no rows. Use the  **AddRow** method to add rows to the new section.

The  **GeometryCount** property returns the number of Geometry sections included in a shape. Use the following code to add a Geometry section to a shape:




```
Shape.AddSection(visSectionFirstComponent + i) 

```

 where 0 <= i < **visSectionLastComponent** - **visSectionFirstComponent** .


- When 0 <= i <  **Shape.GeometryCount** , the new section precedes the present i'th Geometry section.
    
- When  **Shape.GeometryCount** <= i < **visSectionLastComponent** - **visSectionFirstComponent** , the new section is the last section.
    



## Example

The following macro shows how to add a Scratch section to the ShapeSheet of a rectangle. Before running this macro, make sure a drawing page is active in the Visio window.


```vb
 
Public Sub AddSection_Example() 
 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 
 'Get the active page. 
 Set vsoPage = ActivePage 
 
 'Draw a rectangle on the active page. 
 Set vsoShape = vsoPage.DrawRectangle(1, 5, 5, 1) 
 
 'Add a scratch section to the ShapeSheet. 
 vsoShape.AddSection visSectionScratch 
 
End Sub
```


