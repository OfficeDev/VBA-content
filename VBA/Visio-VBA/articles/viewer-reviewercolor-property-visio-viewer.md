---
title: Viewer.ReviewerColor Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ReviewerColor
ms.assetid: 6ec6b962-fc19-1fec-2482-836ab71ece90
ms.date: 06/08/2017
---


# Viewer.ReviewerColor Property (Visio Viewer)

Gets the color of the markup overlay associated with the specified reviewer in the drawing open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **ReviewerColor**( **_ReviewerIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ReviewerIndex|Required| **Long**|The index of the reviewer in the collection of reveiwers.|

### Return Value

 **[OLE_COLOR]**


## Remarks

The property returns a value of data type  **OLE_COLOR** that represents the color of the markup overlay associated with the specified reviewer in Visio Viewer. The **OLE_COLOR** data type is used for properties that return colors.

Valid hexadecimal values for an  **OLE_COLOR** data type in Visio Viewer are of the form _&;Hbbggrr_, where  _bb_ is the blue value, _gg_ the green value, and _rr_ the red value. All three color values range between 00 and FF hexadecimal (255 decimal).

The collection of reviewers is one-based, so the index of the first reviewer in the collection is 1.


## Example

The following code gets the color of the markup overlay of the reviewer at index position 1 in the collection of reviewers in the drawing open in Visio Viewer.


```vb
Debug.Print vsoViewer.ReviewerColor(1)
```


