---
title: Shapes.AddPlaceholder Method (PowerPoint)
keywords: vbapp10.chm543024
f1_keywords:
- vbapp10.chm543024
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddPlaceholder
ms.assetid: 10927d59-1810-2f91-eb52-c42113570ccc
ms.date: 06/08/2017
---


# Shapes.AddPlaceholder Method (PowerPoint)

Restores a previously deleted placeholder on a slide. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the restored placeholder.


## Syntax

 _expression_. **AddPlaceholder**( **_Type_**, **_Left_**, **_Top_**, **_Width_**, **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**[PpPlaceholderType](ppplaceholdertype-enumeration-powerpoint.md)**|The type of placeholder. Placeholders of type  **ppPlaceholderVerticalBody** or **ppPlaceholderVerticalTitle** are found only on slides of layout type **ppLayoutVerticalText**, **ppLayoutClipArtAndVerticalText**, **ppLayoutVerticalTitleAndText**, or **ppLayoutVerticalTitleAndTextOverChart**. You cannot create slides with any of these layouts from the user interface; you must create them programmatically by using the **Add** method or by setting the **Layout** property of an existing slide.|
| _Left_|Optional|**Single**|The position (in points) of the upper-left corner of the placeholder relative to the upper-left corner of the document.|
| _Top_|Optional|**Single**|The position (in points) of the upper-left corner of the placeholder relative to the upper-left corner of the document.|
| _Width_|Optional|**Single**|The width of the placeholder, in points.|
| _Height_|Optional|**Single**|The height of the placeholder, in points.|

### Return Value

Shape


## Remarks

If more than one placeholder of a specified type has been deleted from the slide, the  **AddPlaceholder** method will add them back to the slide, one by one, starting with the placeholder that has the lowest original index number.


## Example

Suppose that slide two in the active presentation originally had a title at the top of the slide that's been deleted, either manually or with the following line of code.


```vb
ActivePresentation.Slides(2).Shapes.Placeholders(1).Delete
```

This example restores the deleted placeholder to slide two.




```vb
Application.ActivePresentation.Slides(2) _
    .Shapes.AddPlaceholder ppPlaceholderTitle
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

