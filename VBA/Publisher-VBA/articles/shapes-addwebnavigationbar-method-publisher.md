---
title: Shapes.AddWebNavigationBar Method (Publisher)
keywords: vbapb10.chm2162736
f1_keywords:
- vbapb10.chm2162736
ms.prod: publisher
api_name:
- Publisher.Shapes.AddWebNavigationBar
ms.assetid: 26e9622c-ea28-b28b-9904-b3a3ccc9341b
ms.date: 06/08/2017
---


# Shapes.AddWebNavigationBar Method (Publisher)

Adds a  **Shape** object of type **pbWebNavigationBar** to the current page of a publication.


## Syntax

 _expression_. **AddWebNavigationBar**( **_Name_**,  **_Left_**,  **_Top_**,  **_Width_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|Required| **String**|The name of the  **WebNavigationBarSet** object to add to the specified **Shape**.|
|Left|Required| **Variant**|The position of the left edge of the shape that represents the Web navigation bar set.|
|Top|Required| **Variant**|The position of the top edge of the shape that represents the Web navigation bar set.|
|Width|Optional| **Variant**|The width of the shape that represents the Web navigation bar set.|

### Return Value

Shape


## Remarks

The  **AddWebNavigationBar** method does not create a Web navigation bar set. It adds an existing set from the **WebNavigationBarSets** collection. Pass the name of the existing Web navigation bar set as the Name parameter.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AddWebNavigationBar** method to add a **WebNavigationBarSet** object to the active document.


```vb
Public Sub AddWebNavigationBarSet_Example() 
 
 Dim pubShape As Publisher.Shape 
 
 ThisDocument.WebNavigationBarSets.AddSet ("NavBar") 
 Set pubShape = ThisDocument.Pages(1).Shapes.AddWebNavigationBar("NavBar", 10, 25) 
 
End Sub
```


