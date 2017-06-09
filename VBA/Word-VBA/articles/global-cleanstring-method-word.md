---
title: Global.CleanString Method (Word)
keywords: vbawd10.chm163119458
f1_keywords:
- vbawd10.chm163119458
ms.prod: word
api_name:
- Word.Global.CleanString
ms.assetid: 787434a2-ff6d-f812-9106-843a69c1cde8
ms.date: 06/08/2017
---


# Global.CleanString Method (Word)

Removes nonprinting characters (character codes 1 ? 29) and special Word characters from the specified string or changes them to spaces (character code 32), as described in the "Remarks" section. Returns the result as a  **String** .


## Syntax

 _expression_ . **CleanString**( **_String_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _String_|Required| **String**|The source string.|

## Remarks

The following characters are converted as described in this table.



|**Character code**|**Description**|
|:-----|:-----|
|7 (beep)|Removed unless preceded by character 13 (paragraph), then converted to character 9 (tab).|
|10 (line feed)|Converted to character 13 (paragraph) unless preceded by character 13, then removed.|
|13 (paragraph)|Unchanged.|
|31 (optional hyphen)|Removed.|
|160 (nonbreaking space)|Converted to character 32 (space).|
|172 (optional hyphen)|Removed.|
|176 (nonbreaking space)|Converted to character 32 (space).|
|182 (paragraph mark)|Removed.|
|183 (bullet)|Converted to character 32 (space).|

## Example

This example removes nonprinting characters from the selected text and inserts the result into a new document.


```vb
Dim strClean As String 
Dim docNew As Document 
 
strClean = Application.CleanString(Selection.Text) 
Set docNew = Documents.Add 
docNew.Content.InsertAfter strClean
```

This example removes nonprinting characters from the selected field code and then displays the result.




```vb
ActiveDocument.ActiveWindow.View.ShowFieldCodes = True 
ActiveDocument.Fields(1).Select 
MsgBox Application.CleanString(Selection.Text)
```


## See also


#### Concepts


[Global Object](global-object-word.md)

