---
title: TextRange.Expand Method (Publisher)
keywords: vbapb10.chm5308421
f1_keywords:
- vbapb10.chm5308421
ms.prod: publisher
api_name:
- Publisher.TextRange.Expand
ms.assetid: 66d8b1a3-5fc4-bed7-94d2-06be6203e1e9
ms.date: 06/08/2017
---


# TextRange.Expand Method (Publisher)

Expands the specified range or selection. Returns or sets a  **Long** that represents the number of specified units added to the range or selection.


## Syntax

 _expression_. **Expand**( **_Unit_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Unit|Required| **PbTextUnit**|The unit by which to expand the range.|

### Return Value

Long


## Remarks

The  **Expand** method moves both endpoints of a range if necessary; to move only one endpoint of a range, use the **[MoveStart](textrange-movestart-method-publisher.md)** or **[MoveEnd](textrange-moveend-method-publisher.md)** method.

The Unit parameter can be one of the  **[PbTextUnit](pbtextunit-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example creates a range that refers to the first word in the first shape of the active publication, formats the font for the word, and then it expands the range to reference the entire first paragraph and formats the font for the whole line.


```vb
Sub ExpandRange() 
 Dim rngText As TextRange 
 
 Set rngText = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Words(Start:=1, Length:=1) 
 With rngText 
 With .Font 
 .Size = 20 
 .Italic = msoTrue 
 End With 
 .Expand Unit:=pbTextUnitLine 
 .Font.Bold = msoTrue 
 End With 
End Sub
```


