---
title: Shape.Type Property (PowerPoint)
keywords: vbapp10.chm547038
f1_keywords:
- vbapp10.chm547038
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Type
ms.assetid: 3a6aa03d-8d93-9a08-ef42-8f128ada7b87
ms.date: 06/08/2017
---


# Shape.Type Property (PowerPoint)

Represents the type of shape or shapes in a range of shapes. Read-only.


## Syntax

 _expression_. **Type**

 _expression_ A variable that represents a **Shape** object.


### Return Value

MsoShapeType


## Remarks

The value of the  **Type** property can be one of these **MsoShapeType** constants.


||
|:-----|
|<strong>msoAutoShape</strong>|
|
<strong>msoCallout</strong>|
|
<strong>msoCanvas</strong>|
|
<strong>msoChart</strong>|
|
<strong>msoComment</strong>|
|
<strong>msoContentApp</strong>|
|
<strong>msoDiagram</strong>|
|
<strong>msoEmbeddedOLEObject</strong>|
|
<strong>msoFormControl</strong>|
|
<strong>msoFreeform</strong>|
|
<strong>msoGroup</strong>|
|
<strong>msoLine</strong>|
|
<strong>msoLinkedOLEObject</strong>|
|
<strong>msoLinkedPicture</strong>|
|
<strong>msoMedia</strong>|
|
<strong>msoOLEControlObject</strong>|
|
<strong>msoPicture</strong>|
|
<strong>msoPlaceholder</strong>|
|
<strong>msoScriptAnchor</strong>|
|
<strong>msoShapeTypeMixed</strong>|
|
<strong>msoTable</strong>|
|
<strong>msoTextBox</strong>|
|
<strong>msoTextEffect</strong>|

## Example

This example loops through all the shapes on all the slides in the active presentation and sets all linked Microsoft Office Excel worksheets to be updated manually.


```vb
For Each sld In ActivePresentation.Slides 
    For Each sh In sld.Shapes 
        If sh.Type = msoLinkedOLEObject Then 
            If sh.OLEFormat.ProgID = "Excel.Sheet" Then 
                sh.LinkFormat.AutoUpdate = ppUpdateOptionManual 
            End If 
        End If 
    Next 
Next
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

