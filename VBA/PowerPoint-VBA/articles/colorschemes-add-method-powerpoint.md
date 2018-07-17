---
title: ColorSchemes.Add Method (PowerPoint)
keywords: vbapp10.chm536004
f1_keywords:
- vbapp10.chm536004
ms.prod: powerpoint
api_name:
- PowerPoint.ColorSchemes.Add
ms.assetid: 1e727a60-0e19-e033-2dc2-c00083263e06
ms.date: 06/08/2017
---


# ColorSchemes.Add Method (PowerPoint)

Adds a color scheme to the collection of available schemes. Returns a  **[ColorScheme](colorscheme-object-powerpoint.md)** object that represents the added color scheme.


## Syntax

 _expression_. **Add**( **_Scheme_** )

 _expression_ A variable that represents a **ColorSchemes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Scheme_|Optional|**ColorScheme**|The color scheme to add. Can be a  **ColorScheme** object from any slide or master or an item in the **ColorSchemes** collection from any open presentation. If you omit this parameter, the first **ColorScheme** object (the first standard color scheme) in the specified presentation's **ColorSchemes** collection is used.|

### Return Value

ColorScheme


## Remarks

The new color scheme is based on the colors used on the specified slide or master or on the colors in the specified color scheme from an open presentation.

The  **ColorSchemes** collection can contain up to 16 color schemes. If you need to add another color scheme and the **ColorSchemes** collection is already full, use the **Delete** method to remove an existing color scheme.

Note that although Microsoft PowerPoint automatically checks whether a color scheme is a duplicate when a user attempts to add it by using the user interface, PowerPoint doesn't check when you attempt to add a color scheme programmatically. Your procedure must do its own checking to avoid adding redundant color schemes.


## See also


#### Concepts


[ColorSchemes Object](colorschemes-object-powerpoint.md)

