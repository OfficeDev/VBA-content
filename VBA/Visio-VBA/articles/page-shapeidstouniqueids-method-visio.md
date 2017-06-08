---
title: Page.ShapeIDsToUniqueIDs Method (Visio)
keywords: vis_sdr.chm10960160
f1_keywords:
- vis_sdr.chm10960160
ms.prod: visio
api_name:
- Visio.Page.ShapeIDsToUniqueIDs
ms.assetid: b89e82db-3c7b-fb73-2f4c-10056c6e7b28
ms.date: 06/08/2017
---


# Page.ShapeIDsToUniqueIDs Method (Visio)

Returns an array of unique IDs of shapes on the page, as specified by their shape IDs.


## Syntax

 _expression_ . **ShapeIDsToUniqueIDs**( **_ShapeIDs()_** , **_UniqueIDArgs_** , **_GUIDs()_** )

 _expression_ An expression that returns a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShapeIDs()_|Required| **Long**|An array of type  **Long** of shape IDs corresponding to a set of shapes on the active drawing page.|
| _UniqueIDArgs_|Required| **VisUniqueIDArgs**|Gets, deletes, or makes the unique ID of a  **Shape** object. See Remarks for possible values.|
| _GUIDs()_|Required| **String**|Out parameter. An empty array that the method fills with unique IDs of type  **String** corresponding to the shapes specified in _ShapeIDs()_|

### Return Value

Nothing


## Remarks

Microsoft Visio identifies shapes by two different IDs: shape IDs and unique IDs.  _Shape IDs_ are numeric and uniquely identify shapes within the scope of an individual drawing page. They are not unique within a wider scope, however.

 _Unique IDs_ are globally unique identifiers (GUIDs). They are unique within the scope of the application.

To convert between shape IDs and unique IDs, you can use two methods of the  **Page** object, **ShapeIDsToUniqueIDs** and **[UniqueIDsToShapeIDs](page-uniqueidstoshapeids-method-visio.md)** .

By default, a shape does not have a unique ID. A shape acquires a unique ID only if you set its  **[Shape.UniqueID](shape-uniqueid-property-visio.md)** property. If a **Shape** object has a unique ID, no other shape in any other document will have the same ID.

The  _UniqueIDArgs_ parameter sets and controls the behavior of the **UniqueID** property for all the shapes in _ShapeIDs()_ . _UniqueIDArgs_ should have one of the following values declared in the Visio type library in **VisUniqueIDArgs** .



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visGetGUID**|0|Returns the unique ID string only if the shape already has a unique ID. Otherwise it returns a zero-length string ("").|
| **visGetOrMakeGUID**|1| Returns the unique ID string of the shape. If the shape does not yet have a unique ID, it assigns one to the shape and returns the new ID.|
| **visDeleteGUID**|2|Deletes the unique ID of a shape and returns a zero-length string ("").|
| **visGetOrMakeGUIDWithUndo**|3|Returns the unique ID string of the shape. If the shape does not already have a unique ID, assigns one to the shape and returns the new ID. Undoable.|
| **visDeleteGUIDWithUndo**|4|Clears the unique ID of a shape and returns a zero-length string (""). Undoable.|

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ShapeIDsToUniqueIDs** method to determine the unique IDs of the shapes on the page passed to the method. It iterates through all the shapes on the active drawing page, using the **Shape.UniqueID** property to get the shape IDs of the shapes, and then passes an array of those IDs to the **ShapeIDsToUniqueIDs** method as the _ShapeIDs()_ parameter to get the unique IDs of the shapes. For the UniqueIDArgs parameter, it passes the value **visGetOrMakeGUID** , telling Visio to create a unique ID for any shape that doesn't already have one. It prints the unique IDs and shape IDs to the **Immediate** window.

Before running this macro, open a Visio drawing and place several shapes on the active drawing page.




```vb
Public Sub ShapeIDsToUniqueIDs_Example()
    Dim vsoShape As Visio.Shape 
    Dim intArrayCounter As Integer 
     
    intShapeCount = ActivePage.Shapes.Count 
     
    ReDim alngShapeIDs(intShapeCount - 1) As Long 
    ReDim astrUniqueIDs(intShapeCount - 1) As String 
     
    intArrayCounter = 0 
         
    For Each vsoShape In ActivePage.Shapes 
        alngShapeIDs(intArrayCounter) = vsoShape.ID 
        Debug.Print alngShapeIDs(intArrayCounter) 
        intArrayCounter = intArrayCounter + 1 
    Next 
     
    ActivePage.ShapeIDsToUniqueIDs alngShapeIDs, visGetOrMakeGUID, astrUniqueIDs 
     
    intArrayCounter = 0 
     
    For intArrayCounter = LBound(astrUniqueIDs) To UBound(astrUniqueIDs) 
        Debug.Print astrUniqueIDs(intArrayCounter) 
    Next 
 
End Sub
```


