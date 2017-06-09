---
title: DocumentWindows.Arrange Method (PowerPoint)
keywords: vbapp10.chm509004
f1_keywords:
- vbapp10.chm509004
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindows.Arrange
ms.assetid: e816fc32-e26f-6a3a-8299-7db24588778f
ms.date: 06/08/2017
---


# DocumentWindows.Arrange Method (PowerPoint)

Arranges all open document windows in the workspace.


## Syntax

 _expression_. **Arrange**( **_arrangeStyle_** )

 _expression_ A variable that represents a **DocumentWindows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _arrangeStyle_|Optional|**[PpArrangeStyle](pparrangestyle-enumeration-powerpoint.md)**|Specifies whether to cascade or tile the windows.|

### Return Value

Nothing


## Example

This example creates a new window and then arranges all open document windows.


```vb
Application.ActiveWindow.NewWindow

Windows.Arrange ppArrangeCascade
```


## See also


#### Concepts


[DocumentWindows Object](documentwindows-object-powerpoint.md)


