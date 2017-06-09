---
title: Viewer.FollowHyperlink Method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.FollowHyperlink
ms.assetid: eafbba6d-6429-744a-facd-e3412916a4bf
ms.date: 06/08/2017
---


# Viewer.FollowHyperlink Method (Visio Viewer)

Follows the hyperlink at the specified index in the specified shape in Microsoft Visio Viewer.


## Syntax

 _expression_. **FollowHyperlink**( **_ShapeIndex_**,  **_HyperlinkIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ShapeIndex|Required| **Long**|The index of the shape that contains the hyperlink.|
|HyperlinkIndex|Required| **Long**|The index of the hyperlink in the collection of hyperlinks in the specified shape.|

### Return Value

Nothing


## Remarks

The collection of hyperlinks is one-based, so the first hyperlink in the collection is at index position 1. If you pass 0 for HyperlinkIndex, Visio Viewer navigates to the default hyperlink for the shape, as set in the  **Hyperlinks** dialog box ( **Insert** menu) in the current Visio document.


## Example

The following code follows the hyperlink in the first index position in the collection of hyperlinks in the first shape on the page in Visio Viewer.


```
vsoViewer.FollowHyperlink 1, 1
```


