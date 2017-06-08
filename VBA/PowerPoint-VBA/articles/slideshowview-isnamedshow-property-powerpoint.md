---
title: SlideShowView.IsNamedShow Property (PowerPoint)
keywords: vbapp10.chm513013
f1_keywords:
- vbapp10.chm513013
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.IsNamedShow
ms.assetid: a68632b2-bff4-9047-f0b8-6acb22a29071
ms.date: 06/08/2017
---


# SlideShowView.IsNamedShow Property (PowerPoint)

Determines whether a custom (named) slide show is displayed in the specified slide show view. Read-only.


## Syntax

 _expression_. **IsNamedShow**

 _expression_ A variable that represents an **SlideShowView** object.


### Return Value

MsoTriState


## Remarks

The value of the  **IsNamedShow** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|A custom (named) slide show is not displayed in the specified slide show view.|
|**msoTrue**| A custom (named) slide show is displayed in the specified slide show view.|

## Example

If the slide show running in slide show window one is a custom slide show, this example displays its name.


```vb
With SlideShowWindows(1).View
    If .IsNamedShow Then
        MsgBox "Now showing in slide show window 1: " _
           &; .SlideShowName
    End If
End With
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

