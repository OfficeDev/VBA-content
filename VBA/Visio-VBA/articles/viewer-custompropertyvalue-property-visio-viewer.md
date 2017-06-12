---
title: Viewer.CustomPropertyValue Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.CustomPropertyValue
ms.assetid: 6e7b87bf-8c2f-3fb6-84a2-a56ee9e59fd7
ms.date: 06/08/2017
---


# Viewer.CustomPropertyValue Property (Visio Viewer)

Gets the value of the shape data item (custom property) at the specified index position for the specified shape in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **CustomPropertyValue**( **_ShapeIndex_**,  **_PropertyIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ShapeIndex|Required| **Long**|The index of the shape that contains the specified shape data item (custom property).|
|PropertyIndex|Required| **Long**|The index of the shape data item (custom property).|

### Return Value

 **String**


## Remarks

In versions of Visio prior to Microsoft Office Visio 2007, shape data items were called custom properties.


## Example

The following code gets the value of the first shape data item assigned to the first shape in the collection of shapes on the current page in Visio Viewer. If the value of the specified custom property is "Hello", Visio Viewer displays a message box and the  **Properties and Settings** dialog box.


```vb
Dim strPropertyValue As String

strPropertyValue = vsoViewer.CustomPropertyValue(1,1)

Debug.Print strPropertyValue

If strPropertyValue = "Hello" Then

    Interaction.MsgBox ("Value is 'Hello'")

    vsoViewer.DisplayPropertyDialog

End If


```


