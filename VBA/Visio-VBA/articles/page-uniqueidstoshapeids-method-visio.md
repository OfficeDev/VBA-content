---
title: Page.UniqueIDsToShapeIDs Method (Visio)
keywords: vis_sdr.chm10960165
f1_keywords:
- vis_sdr.chm10960165
ms.prod: visio
api_name:
- Visio.Page.UniqueIDsToShapeIDs
ms.assetid: 86d0d47c-d356-04ba-51ce-7d682fd165ae
ms.date: 06/08/2017
---


# Page.UniqueIDsToShapeIDs Method (Visio)

Returns an array of shape IDs of shapes on the page, as specifed by their unique IDs.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **UniqueIDsToShapeIDs**( **_GUIDs()_** , **_ShapeIDs()_** )

 _expression_ An expression that returns a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _GUIDs()_|Required| **String**|An array of unique IDs of type  **String** of shapes on the page.|
| _ShapeIDs()_|Required| **Long**|Out parameter. An empty array that the method fills with shape IDs of type  **Long** corresponding to the shapes specified in GUIDs()|

### Return Value

Nothing


## Remarks

Microsoft Visio identifies shapes by two different IDs: shape IDs and unique IDs.  _Shape IDs_ are numeric and uniquely identify shapes within the scope of an individual drawing page. They are not unique within a wider scope, however.

 _Unique IDs_ are globally unique identifiers (GUIDs). They are unique within the scope of the application.

To convert between shape IDs and unique IDs, you can use two methods of the  **Page** object, **[ShapeIDsToUniqueIDs](page-shapeidstouniqueids-method-visio.md)** and **UniqueIDsToShapeIDs** .

By default, a shape does not have a unique ID. A shape acquires a unique ID only if you set its  **[Shape.UniqueID](shape-uniqueid-property-visio.md)** property.

If a  **Shape** object has a unique ID, no other shape in any other document will have the same ID.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **UniqueIDsToShapeIDs** method to determine the shape IDs of the shapes on the page passed to the method as unique IDs. It iterates through all the shapes on the active drawing page, using the **UniqueID** property of each shape to get the unique IDs of the shapes. Then it passes those unique IDs to the **UniqueIDsToShapeIDs** method to return the shape IDs of the shapes. It prints the unique IDs and shape IDs to the **Immediate** window.

Before running this macro, open a Visio drawing and place several shapes on the active drawing page.




```vb
Public Sub UniqueIDsToShapeIDs_Example() 
 
    Dim vsoShape As Visio.Shape 
    Dim intArrayCounter As Integer 
    Dim intShapeCount As Integer 
     
    intShapeCount = ActivePage.Shapes.Count 
     
    ReDim astrUniqueIDs(intShapeCount - 1) As String 
    ReDim alngShapeIDs(intShapeCount - 1) As Long 
     
    intArrayCounter = 0 
     
    For Each vsoShape In ActivePage.Shapes         
        astrUniqueIDs(intArrayCounter) = vsoShape.UniqueID(1) 
        Debug.Print astrUniqueIDs(intArrayCounter) 
        intArrayCounter = intArrayCounter + 1 
    Next 
    
    ActivePage.UniqueIDsToShapeIDs astrUniqueIDs, alngShapeIDs 
     
    intArrayCounter = 0 
 
    For intArrayCounter = LBound(alngShapeIDs) To UBound(alngShapeIDs) 
        Debug.Print alngShapeIDs(intArrayCounter) 
    Next 
 
End Sub
```


