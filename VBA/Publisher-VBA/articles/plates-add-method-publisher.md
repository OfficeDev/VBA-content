---
title: Plates.Add Method (Publisher)
keywords: vbapb10.chm2818052
f1_keywords:
- vbapb10.chm2818052
ms.prod: publisher
api_name:
- Publisher.Plates.Add
ms.assetid: 7fb7b602-8797-e275-4ff7-2e87cf1db11f
ms.date: 06/08/2017
---


# Plates.Add Method (Publisher)

Adds a new color plate to the specified  **Plates** object.


## Syntax

 _expression_. **Add**( **_PlateColor_**)

 _expression_A variable that represents a  **Plates** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PlateColor|Optional| **ColorFormat**| **ColorFormat** object. The color settings to apply to the new plate.|

## Remarks

If the  ** [ColorMode](http://msdn.microsoft.com/library/58befa97-9d9b-9294-18b2-ae10dc87f51c%28Office.15%29.aspx)** property of the specified publication is not **pbColorModeSpot** or **pbColorModeSpotAndProcess**, an error occurs.


## Example

The following example adds a color plate to the active publication if it is a spot-color publication.


```vb
If ActiveDocument.ColorMode = pbColorModeSpot Then 
 ActiveDocument.Plates.Add 
End If
```


