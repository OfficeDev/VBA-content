---
title: Document.FindShapeByWizardTag Method (Publisher)
keywords: vbapb10.chm196690
f1_keywords:
- vbapb10.chm196690
ms.prod: publisher
api_name:
- Publisher.Document.FindShapeByWizardTag
ms.assetid: c6db9ba7-15b0-e8f0-1ed2-08b6e978c948
ms.date: 06/08/2017
---


# Document.FindShapeByWizardTag Method (Publisher)

Returns a  **ShapeRange** object representing one or all of the shapes placed in a publication by a wizard and bearing the specified wizard tag.


## Syntax

 _expression_. **FindShapeByWizardTag**( **_WizardTag_**,  **_Instance_**)

 _expression_A variable that represents a  **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|WizardTag|Required| **PbWizardTag**|Specifies the wizard tag for which to search.|
|Instance|Optional| **Long**|Specifies which instance of a shape with the specified wizard tag is returned. For Instance equal to n, the nth instance of a shape with the specified wizard tag is returned. If no value for Instance is specified, all the shapes with the specified wizard tag are returned.|

### Return Value

ShapeRange


## Remarks

The WizardTag parameter can be one of the  **[PbWizardTag](pbwizardtag-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

The following example finds the second instance of a shape with the wizard tag  **pbWizardDate** and assigns it to a variable.


```vb
Dim shpWizardTag As Shape 
 
Set shpWizardTag = ActiveDocument._ 
 FindShapeByWizardTag(WizardTag:=pbWizardDate, Instance:=2)
```


