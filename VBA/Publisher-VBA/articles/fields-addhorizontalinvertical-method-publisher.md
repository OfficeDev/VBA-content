---
title: Fields.AddHorizontalInVertical Method (Publisher)
keywords: vbapb10.chm6029319
f1_keywords:
- vbapb10.chm6029319
ms.prod: publisher
api_name:
- Publisher.Fields.AddHorizontalInVertical
ms.assetid: 4b451a24-0d79-70d4-4910-2725f1ed0297
ms.date: 06/08/2017
---


# Fields.AddHorizontalInVertical Method (Publisher)

Inserts horizontal text into a stream of vertical text and returns the new horizontal text as a  **Field** object.


## Syntax

 _expression_. **AddHorizontalInVertical**( **_Range_**,  **_Text_**)

 _expression_A variable that represents a  **Fields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Range|Required| **TextRange**|The text range at which to insert the horizontal text.|
|Text|Required| **String**|The text to be horizontally inserted.|

### Return Value

Field


## Example

This example horizontally inserts the text "horizontal test" after the existing vertical text in shape one on page one of the active publication.


```vb
Dim rngTemp As TextRange 
Dim fldTemp As Field 
 
With ActiveDocument.Pages(1).Shapes(1) 
 Set rngTemp = .TextFrame.TextRange.InsertAfter("") 
 
 Set fldTemp = .TextFrame.TextRange.Fields _ 
 .AddHorizontalInVertical(Range:=rngTemp, Text:="horizontal test") 
End With
```


