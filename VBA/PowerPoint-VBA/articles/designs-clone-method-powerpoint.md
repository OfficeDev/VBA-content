---
title: Designs.Clone Method (PowerPoint)
keywords: vbapp10.chm643006
f1_keywords:
- vbapp10.chm643006
ms.prod: powerpoint
api_name:
- PowerPoint.Designs.Clone
ms.assetid: 2365a43f-8adc-ad26-97fc-0376aedf0b80
ms.date: 06/08/2017
---


# Designs.Clone Method (PowerPoint)

Creates a copy of a  **[Design](design-object-powerpoint.md)** object.


## Syntax

 _expression_. **Clone**( **_pOriginal_**, **_Index_** )

 _expression_ A variable that represents a **Designs** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pOriginal_|Required|**Design**|**Design** object. The original design.|
| _Index_|Optional|**Long**|The index location in the  **Designs** collection into which the design will be copied. If Index is omitted, the cloned design is added to the end of the **Designs** collection.|

### Return Value

Design


## Example

This example creates a design and clones the newly created design.


```vb
Sub CloneDesign()

    Dim dsnDesign1 As Design
    Dim dsnDesign2

    Set dsnDesign1 = ActivePresentation.Designs _
        .Add(designName:="Design1")

    Set dsnDesign2 = ActivePresentation.Designs _
        .Clone(pOriginal:=dsnDesign1, Index:=1)

End Sub
```


## See also


#### Concepts


[Designs Object](designs-object-powerpoint.md)

