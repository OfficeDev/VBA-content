---
title: ShapeRange.WizardTagInstance Property (Publisher)
keywords: vbapb10.chm2293873
f1_keywords:
- vbapb10.chm2293873
ms.prod: publisher
api_name:
- Publisher.ShapeRange.WizardTagInstance
ms.assetid: 07d1c4c8-8efb-b029-2dba-37fef435cc8b
ms.date: 06/08/2017
---


# ShapeRange.WizardTagInstance Property (Publisher)

Returns or sets a  **Long** indicating the instance of the specified shape compared with other shapes having the same wizard tag. Read/write.


## Syntax

 _expression_. **WizardTagInstance**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks

The combination of the  **WizardTagInstance** property and the **[WizardTag](shaperange-wizardtag-property-publisher.md)** property uniquely defines every shape in a publication.


## Example

The following example displays the wizard tag and wizard tag instance information for all the shapes on page one of the active publication.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop 
 Debug.Print "Shape: " &; .Name 
 Debug.Print " Wizard tag: " &; .WizardTag 
 Debug.Print " Wizard tag instance: " _ 
 &; .WizardTagInstance 
 End With 
Next shpLoop
```


