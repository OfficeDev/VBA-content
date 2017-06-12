---
title: Frameset.AddNewFrame Method (Word)
keywords: vbawd10.chm165806130
f1_keywords:
- vbawd10.chm165806130
ms.prod: word
api_name:
- Word.Frameset.AddNewFrame
ms.assetid: 81366e66-ae4e-24ce-d7ca-ae6f9273f745
ms.date: 06/08/2017
---


# Frameset.AddNewFrame Method (Word)

Adds a new frame to a frames page.


## Syntax

 _expression_ . **AddNewFrame**( **_Where_** )

 _expression_ Required. A variable that represents a **[Frameset](frameset-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Where_|Required| **WdFramesetNewFrameLocation**|Sets the location where the new frame is to be added in relation to the specified frame.|

## Remarks

For more information on creating frames pages, see [Creating frames pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example adds a new frame to the immediate right of the specified frame.


```vb
ActiveDocument.ActiveWindow.ActivePane.Frameset _ 
 .AddNewFrame wdFramesetNewFrameRight
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

