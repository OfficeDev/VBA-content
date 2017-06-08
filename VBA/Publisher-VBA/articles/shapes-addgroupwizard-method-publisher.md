---
title: Shapes.AddGroupWizard Method (Publisher)
keywords: vbapb10.chm2162727
f1_keywords:
- vbapb10.chm2162727
ms.prod: publisher
api_name:
- Publisher.Shapes.AddGroupWizard
ms.assetid: 5a84f055-7f30-0757-f507-40ee34b214f4
ms.date: 06/08/2017
---


# Shapes.AddGroupWizard Method (Publisher)

Adds a  **Shape** object representing a Design Gallery object to the publication.


## Syntax

 _expression_. **AddGroupWizard**( **_Wizard_**,  **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**,  **_Design_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Wizard|Required| **PbWizardGroup**|The type of Design Gallery object to add to the publication.|
|Left|Required| **Variant**|The position of the Design Gallery object's left edge relative to the left edge of the page, measured in points.|
|Top|Required| **Variant**|The position of the Design Gallery object's top edge relative to the top edge of the page, measured in points.|
|Width|Optional| **Variant**|The width of the new Design Gallery object.|
|Height|Optional| **Variant**|The height of the new Design Gallery object.|
|Design|Optional| **Long**|The design of the object to be added.|

### Return Value

Shape


## Remarks

The Wizard parameter can be one of the  **[PbWizardGroup](pbwizardgroup-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example adds a Web table of contents to the active publication.


```vb
ActiveDocument.Pages(1).Shapes _ 
 .AddGroupWizard Wizard:=pbWizardGroupTableOfContents, _ 
 Left:=100, Top:=100
```


