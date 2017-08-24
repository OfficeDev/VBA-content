---
title: Shape.WizardTag Property (Publisher)
keywords: vbapb10.chm2228324
f1_keywords:
- vbapb10.chm2228324
ms.prod: publisher
api_name:
- Publisher.Shape.WizardTag
ms.assetid: b93bbdf9-6ce7-3ba6-566a-b11f8044fbda
ms.date: 06/08/2017
---


# Shape.WizardTag Property (Publisher)

Returns or sets a  **PbWizardTag**constant indicating the function of a specified shape with respect to its publication design. Read/write.


## Syntax

 _expression_. **WizardTag**

 _expression_A variable that represents a  **Shape** object.


## Remarks

The  **WizardTag** property value can be one of the **[PbWizardTag](pbwizardtag-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.

The combination of the  **[WizardTagInstance](shape-wizardtaginstance-property-publisher.md)** property and the **WizardTag** property uniquely defines every shape in a publication.


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


