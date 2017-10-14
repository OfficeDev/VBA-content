---
title: AutoTextEntries.AppendToSpike Method (Word)
keywords: vbawd10.chm154599526
f1_keywords:
- vbawd10.chm154599526
ms.prod: word
api_name:
- Word.AutoTextEntries.AppendToSpike
ms.assetid: c54857c4-1a4b-34fc-8510-592276bd1753
ms.date: 06/08/2017
---


# AutoTextEntries.AppendToSpike Method (Word)

Deletes the specified range and adds the contents of the range to the Spike (a built-in AutoText entry). This method returns the Spike as an  **AutoTextEntry** object.


## Syntax

 _expression_ . **AppendToSpike**( **_Range_** )

 _expression_ Required. A variable that represents an **[AutoTextEntries](autotextentries-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range that's deleted and appended to the Spike.|

### Return Value

AutoTextEntry


## Remarks

The  **AppendToSpike** method is only valid for the **AutoTextEntries** collection in the Normal template.


## Example

This example deletes the selection and adds its contents to the Spike in the Normal template.


```vb
If Len(Selection.Range.Text) > 1 Then 
 NormalTemplate.AutoTextEntries.AppendToSpike _ 
 Range:=Selection.Range 
End If
```

This example clears the Spike and adds the first and third words in the active document to the Spike in the Normal template. The contents of the Spike are then inserted at the insertion point.




```vb
Dim atEntry As AutoTextEntry 
Selection.Collapse Direction:=wdCollapseStart 
For Each atEntry In NormalTemplate.AutoTextEntries 
 If atEntry.Name = "Spike" Then atEntry.Delete 
Next atEntry 
With NormalTemplate.AutoTextEntries 
 .AppendToSpike Range:=ActiveDocument.Words(3) 
 .AppendToSpike Range:=ActiveDocument.Words(1) 
 .Item("Spike").Insert Where:=Selection.Range 
End With
```


## See also


#### Concepts


[AutoTextEntries Collection Object](autotextentries-object-word.md)

