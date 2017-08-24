---
title: TextRange.Lines Method (Publisher)
keywords: vbapb10.chm5308455
f1_keywords:
- vbapb10.chm5308455
ms.prod: publisher
api_name:
- Publisher.TextRange.Lines
ms.assetid: 56862090-b2ff-403b-d016-e37108d5ccc1
ms.date: 06/08/2017
---


# TextRange.Lines Method (Publisher)

Returns a  **[TextRange](textrange-object-publisher.md)** object that represents the specified lines.


## Syntax

 _expression_. **Lines**( **_Start_**,  **_Length_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Start|Required| **Long**|The first line in the returned range.|
|Length|Optional| **Long**|The number of lines to be returned. Default is 1.|

### Return Value

TextRange


## Remarks

If  **_Start_** is greater than the number of lines in the specified text, the returned range starts with the last line in the specified range.

If  **_Length_** is greater than the number of lines from the specified starting line to the end of the text, the returned range contains all those lines.


## Example

This example replaces the first three lines of the first shape on the first page with the specified string.


```vb
Sub ReplaceLines() 
 Dim rngText As TextRange 
 Set rngText = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Lines(Start:=1, Length:=3) 
 
 rngText.Text = "This is replacement text." &; vbCrLf 
 
End Sub
```


