---
title: CalloutFormat.CustomDrop Method (Publisher)
keywords: vbapb10.chm2490385
f1_keywords:
- vbapb10.chm2490385
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.CustomDrop
ms.assetid: 65fc7309-acd0-5bdd-6bb0-1b6c41968775
ms.date: 06/08/2017
---


# CalloutFormat.CustomDrop Method (Publisher)

Sets the vertical distance from the edge of the text bounding box to the place where the callout line attaches to the text box.


## Syntax

 _expression_. **CustomDrop**( **_Drop_**)

 _expression_A variable that represents a  **CalloutFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Drop|Required| **Variant**|The drop distance. Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|

## Remarks

The drop distance is normally measured from the top of the text box. However, if the  **[AutoAttach](calloutformat-autoattach-property-publisher.md)** property is set to **True** and the text box is to the left of the origin of the callout line (the place to which the callout points), the drop distance is measured from the bottom of the text box.


## Example

This example sets the custom drop distance to 14 points, and specifies that the drop distance always be measured from the top. For the example to work, the third shape in the active publication must be a callout.


```vb
With ActiveDocument.Pages(1).Shapes(3).Callout 
 .CustomDrop Drop:=14 
 .AutoAttach = False 
End With 

```


