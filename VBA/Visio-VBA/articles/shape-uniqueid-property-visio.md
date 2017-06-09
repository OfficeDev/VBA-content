---
title: Shape.UniqueID Property (Visio)
keywords: vis_sdr.chm11214615
f1_keywords:
- vis_sdr.chm11214615
ms.prod: visio
api_name:
- Visio.Shape.UniqueID
ms.assetid: a82e1175-4536-8919-6531-593d57c3b2f5
ms.date: 06/08/2017
---


# Shape.UniqueID Property (Visio)

Gets, deletes, or makes the GUID that uniquely identifies the shape within the scope of the application. Read-only.


## Syntax

 _expression_ . **UniqueID**( **_fUniqueID_** )

 _expression_ An expression that returns a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fUniqueID_|Required| **Integer**|Gets, deletes, or makes the unique ID of a  **Shape** object. See Remarks for possible values.|

### Return Value

String


## Remarks

Microsoft Visio identifies shapes by two different IDs: shape IDs and unique IDs.  _Shape IDs_ are numeric and uniquely identify shapes within the scope of an individual drawing page. They are not unique within a wider scope, however.

 _Unique IDs_ are GUIDs. They are unique within the scope of the application.

To convert between shape IDs and unique IDs, you can use two methods of the  **Page** object, **[ShapeIDsToUniqueIDs](page-shapeidstouniqueids-method-visio.md)** and **[UniqueIDsToShapeIDs](page-uniqueidstoshapeids-method-visio.md)** .

By default, a shape does not have a unique ID. A shape acquires a unique ID only if you set its  **UniqueID** property.

If a  **Shape** object has a unique ID, no other shape in any other document will have the same ID.

The  _fUniqueID_ parameter controls the behavior of the **UniqueID** property. It should have one of the following values declared in the Visio type library in **VisUniqueIDArgs** .



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visGetGUID**|0|Returns the unique ID string only if the shape already has a unique ID. Otherwise it returns a zero-length string ("").|
| **visGetOrMakeGUID**|1| Returns the unique ID string of the shape. If the shape does not yet have a unique ID, it assigns one to the shape and returns the new ID.|
| **visDeleteGUID**|2|Deletes the unique ID of a shape and returns a zero-length string ("").|
| **visGetOrMakeGUIDWithUndo**|3|Returns the unique ID string of the shape. If the shape does not already have a unique ID, assigns one to the shape and returns the new ID. Undoable.|
| **visDeleteGUIDWithUndo**|4|Clears the unique ID of a shape and returns a zero-length string (""). Undoable.|
To get a shape if you know its unique ID, use  **Shapes.Item** ( _UniqueIDString_).

For example, you can use the following code:




```vb
Dim vsoShape As Visio.Shape 
Set vsoShape = Visio.ActivePage.Shapes.Item("{2287DC42-B167-11CE-88E9-0020AFDDD917}") 

```

Alternatively, you can use the following code, which adds the letter "U" before the string to identify it as a unique ID:




```vb
Dim vsoShape As Visio.Shape 
Set vsoShape = Visio.ActivePage.Shapes.Item("U{2287DC42-B167-11CE-88E9-0020AFDDD917}") 

```


