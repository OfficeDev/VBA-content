---
title: Global.Repeat Method (Word)
keywords: vbawd10.chm163119409
f1_keywords:
- vbawd10.chm163119409
ms.prod: word
api_name:
- Word.Global.Repeat
ms.assetid: 23e2e300-cc01-cd9d-f761-0113a07267bd
ms.date: 06/08/2017
---


# Global.Repeat Method (Word)

Repeats the most recent editing action one or more times. Returns  **True** if the commands were repeated successfully.


## Syntax

 _expression_ . **Repeat**( **_Times_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Times_|Optional| **Variant**|The number of times you want to repeat the last command.|

### Return Value

Boolean


## Remarks

Using this method is the equivalent to using the  **Repeat** command on the **Edit** menu.


## Example

This example inserts the text "Hello" followed by two paragraphs (the second typing action is repeated once).


```
Selection.TypeText "Hello" 
Selection.TypeParagraph 
Repeat
```

This example repeats the last command three times (if it can be repeated).




```vb
On Error Resume Next 
If Repeat(3) = True Then StatusBar = "Action repeated"
```


## See also


#### Concepts


[Global Object](global-object-word.md)

