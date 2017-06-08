---
title: HTMLDivision.HTMLDivisionParent Method (Word)
keywords: vbawd10.chm166133768
f1_keywords:
- vbawd10.chm166133768
ms.prod: word
api_name:
- Word.HTMLDivision.HTMLDivisionParent
ms.assetid: fee0eaa1-3985-f4fc-4adb-14f0defd9084
ms.date: 06/08/2017
---


# HTMLDivision.HTMLDivisionParent Method (Word)

Returns an  **HTMLDivision** object that represents a parent division of the current HTML division.


## Syntax

 _expression_ . **HTMLDivisionParent**( **_LevelsUp_** )

 _expression_ Required. A variable that represents an **[HTMLDivision](htmldivision-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LevelsUp_|Optional| **Long**|The number of parent divisions to count back to return the desired division. If the LevelsUp argument is omitted, the HTML division returned is one level up from the current HTML division.|

### Return Value

HTMLDivision


## Example

This example formats the borders for two HTML divisions in the active document. This example assumes that the active document is an HTML document with at least two divisions.


```vb
Sub FormatHTMLDivisions() 
 With ActiveDocument.HTMLDivisions(1) 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderLeft) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .Borders(wdBorderRight) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .HTMLDivisionParent 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderTop) 
 .Color = wdColorBlack 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderBottom) 
 .Color = wdColorBlack 
 .LineStyle = wdLineStyleDot 
 End With 
 End With 
 End With 
 End With 
End Sub
```


## See also


#### Concepts


[HTMLDivision Object](htmldivision-object-word.md)

