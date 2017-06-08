---
title: Shapes.ItemFromUniqueID Property (Visio)
keywords: vis_sdr.chm11362485
f1_keywords:
- vis_sdr.chm11362485
ms.prod: visio
api_name:
- Visio.Shapes.ItemFromUniqueID
ms.assetid: 94175764-d65d-9511-4073-864ff89f573c
ms.date: 06/08/2017
---


# Shapes.ItemFromUniqueID Property (Visio)

Returns the  **[Shape](shape-object-visio.md)** object that matches the specified **[UniqueID](shape-uniqueid-property-visio.md)** property value. Read-only.


## Syntax

 _expression_ . **ItemFromUniqueID**( **_UniqueID_** )

 _expression_ A variable that represents a **[Shapes](shapes-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UniqueID_|Required| **String**|The unique ID of a  **Shape** object.|

### Return Value

 **Shape**


## Remarks

Microsoft Visio identifies shapes by two different IDs: shape IDs and unique IDs. Shape IDs are numeric and uniquely identify shapes within the scope of an individual drawing page or master. They are not unique within the scope of the drawing, however.

Unique IDs are GUIDs. They are unique within the scope of the document.

To convert between shape IDs and unique IDs, you can use two methods of the  **[Page](page-object-visio.md)** object, **[ShapeIDsToUniqueIDs](page-shapeidstouniqueids-method-visio.md)** and **[UniqueIDsToShapeIDs](page-uniqueidstoshapeids-method-visio.md)** .

By default, a shape does not have a unique ID. A shape acquires a unique ID only if you get its read-only  **UniqueID** property value by calling the property on the shape, passing it the **visGetOrMake** constant from the **[VisUniqueIDArgs](visuniqueidargs-enumeration-visio.md)** enumeration.

If a  **Shape** object has a unique ID, no other shape in the same document will have the same ID.


