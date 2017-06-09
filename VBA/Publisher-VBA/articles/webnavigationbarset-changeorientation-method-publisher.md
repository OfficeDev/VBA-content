---
title: WebNavigationBarSet.ChangeOrientation Method (Publisher)
keywords: vbapb10.chm8519699
f1_keywords:
- vbapb10.chm8519699
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.ChangeOrientation
ms.assetid: bce05e9c-5b4a-f5a2-33a9-b40d4e05664f
ms.date: 06/08/2017
---


# WebNavigationBarSet.ChangeOrientation Method (Publisher)

Sets a  **PbNavBarOrientation** constant that represents the alignment of the navigation bar; vertical or horizontal.


## Syntax

 _expression_. **ChangeOrientation**( **_Orientation_**)

 _expression_A variable that represents a  **WebNavigationBarSet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Orientation|Required| **PbNavBarOrientation**|

## Remarks

The Orientation parameter can be one of the following  **PbNavBarOrientation** constants declared in the Microsoft Publisher type library.



| **pbNavBarOrientHorizontal**|
| **pbNavBarOrientVertical**|

## Example

The following example sets an object variable to the first Web navigation bar set in the active document, adds it to every page, changes the orientation to horizontal, sets the horizontal alignment to center, and then sets the horizontal button count to 4.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets(1) 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ChangeOrientation pbNavBarOrientHorizontal 
 .HorizontalAlignment = pbnbAlignCenter 
 .HorizontalButtonCount = 4 
End With
```


