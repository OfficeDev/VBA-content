---
title: OMath Object (Word)
keywords: vbawd10.chm2691
f1_keywords:
- vbawd10.chm2691
ms.prod: word
api_name:
- Word.OMath
ms.assetid: 82f2f81b-e2d5-140f-bdcc-8b52b821b24d
ms.date: 06/08/2017
---


# OMath Object (Word)

Represents an equation.  **OMath** objects are members of the **OMaths** collection.


## Remarks

Use the  **Add** method of the **OMaths** collection to create an equation and add it to a document, selection, or range. The following example creates an equation and uses the **BuildUp** method to convert the equation to professional format.


```
Dim objRange As Range 
Dim objEq As OMath 
 
Set objRange = Selection.Range 
objRange.Text = "Celsius = (5/9)(Fahrenheit - 32)" 
Set objRange = Selection.OMaths.Add(objRange) 
Set objEq = objRange.OMaths(1) 
objEq.BuildUp
```


## Methods



|**Name**|
|:-----|
|[BuildUp](omath-buildup-method-word.md)|
|[ConvertToLiteralText](omath-converttoliteraltext-method-word.md)|
|[ConvertToMathText](omath-converttomathtext-method-word.md)|
|[ConvertToNormalText](omath-converttonormaltext-method-word.md)|
|[Linearize](omath-linearize-method-word.md)|
|[Remove](omath-remove-method-word.md)|

## Properties



|**Name**|
|:-----|
|[AlignPoint](omath-alignpoint-property-word.md)|
|[Application](omath-application-property-word.md)|
|[ArgIndex](omath-argindex-property-word.md)|
|[ArgSize](omath-argsize-property-word.md)|
|[Breaks](omath-breaks-property-word.md)|
|[Creator](omath-creator-property-word.md)|
|[Functions](omath-functions-property-word.md)|
|[Justification](omath-justification-property-word.md)|
|[NestingLevel](omath-nestinglevel-property-word.md)|
|[Parent](omath-parent-property-word.md)|
|[ParentArg](omath-parentarg-property-word.md)|
|[ParentCol](omath-parentcol-property-word.md)|
|[ParentFunction](omath-parentfunction-property-word.md)|
|[ParentOMath](omath-parentomath-property-word.md)|
|[ParentRow](omath-parentrow-property-word.md)|
|[Range](omath-range-property-word.md)|
|[Type](omath-type-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
