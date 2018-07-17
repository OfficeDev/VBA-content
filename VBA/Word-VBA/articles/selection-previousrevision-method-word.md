---
title: Selection.PreviousRevision Method (Word)
keywords: vbawd10.chm158663188
f1_keywords:
- vbawd10.chm158663188
ms.prod: word
api_name:
- Word.Selection.PreviousRevision
ms.assetid: e516037f-047d-5cd2-19b4-3b7870a14b5a
ms.date: 06/08/2017
---


# Selection.PreviousRevision Method (Word)

Locates and returns the previous tracked change as a  **Revision** object.


## Syntax

 _expression_ . **PreviousRevision**( **_Wrap_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wrap_|Optional| **Variant**| **True** to continue searching for a revision at the end of the document when the beginning of the document is reached. The default value is **False** .|

### Return Value

Revision


## Example

This example selects the last tracked change in the first section in the active document and displays the date and time of the change.


```vb
Selection.EndOf Unit:=wdStory, Extend:=wdMove 
Set myRev = Selection.PreviousRevision 
If Not (myRev Is Nothing) Then MsgBox myRev.Date
```

This example rejects the previous tracked change found if the change type is deleted or inserted text. If the tracked change is a style change, the change is accepted.




```vb
Set myRev = Selection.PreviousRevision(Wrap:=True) 
If Not (myRev Is Nothing) Then 
 Select Case myRev.Type 
 Case wdRevisionDelete 
 myRev.Reject 
 Case wdRevisionInsert 
 myRev.Reject 
 Case wdRevisionStyle 
 myRev.Accept 
 End Select 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

