---
title: Frameset.ChildFramesetItem Property (Word)
keywords: vbawd10.chm165806086
f1_keywords:
- vbawd10.chm165806086
ms.prod: word
api_name:
- Word.Frameset.ChildFramesetItem
ms.assetid: a0210de1-5556-0c20-a694-a6892dc7eddf
ms.date: 06/08/2017
---


# Frameset.ChildFramesetItem Property (Word)

Returns the  **Frameset** object that represents the child **Frameset** object specified by the Index argument. Read-only.


## Syntax

 _expression_ . **ChildFramesetItem**( **_Index_** )

 _expression_ An expression that returns a **[Frameset](frameset-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the specified frame.|

## Remarks

This property applies only to  **Frameset** objects of type **wdFramesetTypeFrameset** .

For more information on creating frames pages, see [Creating Frames Pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example sets the name of the third child frame of the specified frame to "BottomFrame".


```vb
ActiveWindow.Document.Frameset _ 
 .ChildFramesetItem(3).FrameName = "BottomFrame"
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

