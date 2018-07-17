---
title: Selection.GoTo Method (Word)
keywords: vbawd10.chm158662829
f1_keywords:
- vbawd10.chm158662829
ms.prod: word
api_name:
- Word.Selection.GoTo
ms.assetid: 7a69e581-4047-ae62-e112-97fe2c2633bb
ms.date: 06/08/2017
---

# Selection.GoTo Method (Word)

Moves the insertion point to the character position immediately preceding the specified item, and returns a  **[Range](range-object-word.md)** object (except for the **wdGoToGrammaticalError** , **wdGoToProofreadingError** , or **wdGoToSpellingError** constant).

## Syntax

 _expression_ . **GoTo**( **_What_** , **_Which_** , **_Count_** , **_Name_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.

### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _What_|Optional| **Variant**|The kind of item to which the range or selection is moved. Can be one of the  **[WdGoToItem](wdgotoitem-enumeration-word.md)** constants.|
| _Which_|Optional| **Variant**|The item to which the range or selection is moved. Can be one of the  **[WdGoToDirection](wdgotodirection-enumeration-word.md)** constants.|
| _Count_|Optional| **Variant**|The number of the item in the document. The default value is 1. Only positive values are valid. To specify an item that precedes the range or selection, use  **wdGoToPrevious** as the Which argument and specify a Count value.|
| _Name_|Optional| **Variant**|If the What argument is  **wdGoToBookmark** , **wdGoToComment** , **wdGoToField** , or **wdGoToObject** , this argument specifies a name.|

### Return Value

The [Range](range-object-word.md) that is now selected.

## Remarks

When you use the  **GoTo** method with the **wdGoToGrammaticalError** , **wdGoToProofreadingError** , or **wdGoToSpellingError** constant, the **Range** object that is returned includes any grammar error text or spelling error text.

## Examples

The following examples are functionally equivalent; they both move the selection to the first heading in the document.

```
Selection.GoTo What:=wdGoToHeading, Which:=wdGoToFirst
Selection.GoTo What:=wdGoToHeading, Which:=wdGoToAbsolute, Count:=1
```

The following example moves the selection to the fourth line in the document.

```
Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=4
```

The following example moves the selection up two lines.

```
Selection.GoTo What:=wdGoToLine, Which:=wdGoToPrevious, Count:=2
```

The following example moves to the next DATE field.

```
Selection.GoTo What:=wdGoToField, Name:="Date"
```

This example moves the selection to the first cell in the next table.

```
Selection.GoTo What:=wdGoToTable, Which:=wdGoToNext
```

This example moves the insertion point just before the fifth endnote reference mark in the active document.

```vb
If ActiveDocument.Endnotes.Count >= 5 Then
 Selection.GoTo What:=wdGoToEndnote, _
 Which:=wdGoToAbsolute, Count:=5
End If
```

This example moves the selection down four lines.

```
Selection.GoTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=4
```

This example moves the selection back two pages.

```
Selection.GoTo What:=wdGoToPage, Which:=wdGoToPrevious, Count:=2
```

## See also

#### Concepts

[Selection Object](selection-object-word.md)

