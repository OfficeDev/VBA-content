---
title: ReaderSpread.Pages Property (Publisher)
keywords: vbapb10.chm524293
f1_keywords:
- vbapb10.chm524293
ms.prod: publisher
api_name:
- Publisher.ReaderSpread.Pages
ms.assetid: 181c37b2-ed3f-826a-5718-ae6aff120eb3
ms.date: 06/08/2017
---


# ReaderSpread.Pages Property (Publisher)

Returns a  **[Page](page-object-publisher.md)** object representing one of the pages that compose the specified reader spread.


## Syntax

 _expression_. **Pages**( **_Index_**)

 _expression_A variable that represents a  **ReaderSpread** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The page from the reader spread to return. Can be either 1 or 2.|

## Remarks

A reader spread will consist of only one or two pages, which is why the valid values for the  **Index** argument are 1 or 2.


## Example

The following example checks the reader spread of the fifth page in the active publication to see if it contains more than one page. If it does, the example reports the page number of the second page in the spread.


```vb
Dim pageTemp As Page 
 
With ActiveDocument.Pages(5).ReaderSpread 
 If .PageCount > 1 Then 
 Set pageTemp = .Pages(Index:=2) 
 MsgBox "The page number of the second page " _ 
 &; "in the spread is " &; pageTemp.PageNumber 
 Else 
 MsgBox "The spread has only one page." 
 End If 
End With
```


