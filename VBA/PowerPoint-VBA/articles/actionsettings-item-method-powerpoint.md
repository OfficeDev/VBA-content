---
title: ActionSettings.Item Method (PowerPoint)
keywords: vbapp10.chm566003
f1_keywords:
- vbapp10.chm566003
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSettings.Item
ms.assetid: 88e0b49b-0518-559b-243f-c369c09ab3fe
ms.date: 06/08/2017
---


# ActionSettings.Item Method (PowerPoint)

Returns a single action setting from the specified  **ActionSettings** collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ A variable that represents an **ActionSettings** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**PpMouseActivation**|The action setting for a  **MouseClick** or **MouseOver** event.|

### Return Value

ActionSetting


## Remarks

The  _Index_ parameter value can be one of these **PpMouseActivation** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**ppMouseClick**|The action setting for when the user clicks the shape.|
|**ppMouseOver**|The action setting for when the mouse pointer is positioned over the specified shape.|
The  **Item** method is the default member for a collection. For example, the following two lines of code are equivalent:




```vb
ActivePresentation.Slides.Item(1)
```




```vb
ActivePresentation.Slides(1)
```

For more information about returning a single member of a collection, see [Returning an Object from a Collection](return-objects-from-collections.md).


## Example

This example sets shape three on slide one to play the sound of applause and uses the  **[AnimateAction](actionsetting-animateaction-property-powerpoint.md)** property to specify that the shape's color is to be momentarily inverted when the shape is clicked during a slide show.


```vb
With ActivePresentation.Slides.Item(1).Shapes _
        .Item(3).ActionSettings.Item(ppMouseClick)
    .SoundEffect.Name = "applause"
    .AnimateAction = True
End With
```


## See also


#### Concepts


[ActionSettings Object](actionsettings-object-powerpoint.md)

