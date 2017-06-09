---
title: Shape.ConnectedShapes Method (Visio)
keywords: vis_sdr.chm11262240
f1_keywords:
- vis_sdr.chm11262240
ms.prod: visio
api_name:
- Visio.Shape.ConnectedShapes
ms.assetid: 7f5a0ac9-d0a7-d9fe-9ecb-8e8070ab5951
ms.date: 06/08/2017
---


# Shape.ConnectedShapes Method (Visio)

Returns an array that contains the identifiers (IDs) of the shapes that are connected to the shape.


## Syntax

 _expression_ . **ConnectedShapes**( **_Flags_** , **_CategoryFilter_** )

 _expression_ A variable that represents a **[Shape](shape-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Flags_|Required| **[VisConnectedShapesFlags](visconnectedshapesflags-enumeration-visio.md)**|Filters the array of returned shape IDs by the directionality of the connectors. See Remarks for possible values.|
| _CategoryFilter_|Required| **String**|Filters the array of returned shape IDs by limiting it to the IDs of shapes that match the specified category.|

### Return Value

 **Long()**


## Remarks

The  _Flags_ value must be one of the following **VisConnectedShapesFlags** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visConnectedShapesAllNodes**|0|Return IDs of shapes that are associated with both incoming and outgoing connections.|
| **visConnectedShapesIncomingNodes**|1|Return IDs of shapes that are associated with incoming connections.|
| **visConnectedShapesOutgoingNodes**|2|Return IDs of shapes that are associated with outgoing connections.|
Categories are user-defined strings that you can use to categorize shapes and thereby to restrict membership in a container. You can define categories in the User.msvShapeCategories cell in the ShapeSheet for a shape. You can define multiple categories for a shape by separating the categories with semi-colons.

If the source object is a 1-D shape or part of a master, the  **ConnectedShapes** method returns an Invalid Source error.

If no qualifying connected shapes exist, the  **ConnectedShapes** method returns an empty array.


## Examples

The following Visual Basic for Applications (VBA) macro shows how to use the  **ConnectedShapes** method to find the names of all the shapes at the other end of outgoing connections from a selected shape.

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png)[Fred Diggs](http://www.visiozone.com)




```vb
Public Sub ConnectedShapes_Outgoing_Example()
' Get the shapes that are connected to the selected shape
' by outgoing connectors.
    Dim vsoShape As Visio.Shape
    Dim lngShapeIDs() As Long
    Dim intCount As Integer

    If ActiveWindow.Selection.Count = 0 Then
        MsgBox ("Please select a shape that has connections")
        Exit Sub
    Else
        Set vsoShape = ActiveWindow.Selection(1)
    End If

    lngShapeIDs = vsoShape.ConnectedShapes _
      (visConnectedShapesOutgoingNodes, "")
    Debug.Print "Shapes at the end of outgoing connectors:"
    For intCount = 0 To UBound(lngShapeIDs)
        Debug.Print ActivePage.Shapes(lngShapeIDs(intCount)).Name
    Next
End Sub
```

The following VBA macro shows how to use the  **ConnectedShapes** method to find the names of all the shapes at the other end of incoming connections to a selected shape.

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png)[Fred Diggs](http://www.visiozone.com)




```vb
Public Sub ConnectedShapes_Incoming_Example()
' Get the shapes that are at the other end of 
' incoming connections to a selected shape
    Dim vsoShape As Visio.Shape
    Dim lngShapeIDs() As Long
    Dim intCount As Integer

    If ActiveWindow.Selection.Count = 0 Then
        MsgBox ("Please select a shape that has connections.")
        Exit Sub
    Else
        Set vsoShape = ActiveWindow.Selection(1)
    End If

    lngShapeIDs = vsoShape.ConnectedShapes _
      (visConnectedShapesIncomingNodes, "")
    Debug.Print "Shapes that are at the other end of incoming connections:"
    For intCount = 0 To UBound(lngShapeIDs)
        Debug.Print ActivePage.Shapes(lngShapeIDs(intCount)).Name
    Next
End Sub
```


