---
title: Selection.NextRevision Method (Word)
keywords: vbawd10.chm158663187
f1_keywords:
- vbawd10.chm158663187
ms.prod: word
api_name:
- Word.Selection.NextRevision
ms.assetid: 990e3c20-9991-b2cb-aa3b-e64ae8936b34
ms.date: 06/08/2017
---


# Selection.NextRevision Method (Word)

Locates and returns the next tracked change as a  **Revision** object.


## Syntax

 _expression_ . **NextRevision**( **_Wrap_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wrap_|Optional| **Variant**| **True** to continue searching for a revision at the beginning of the document when the end of the document is reached. The default value is **False** .|

### Return Value

Revision


## Remarks

The changed text becomes the current selection. Use the properties of the resulting  **Revision** object to see what type of change it is, who made it, and so forth. Use the methods of the **Revision** object to accept or reject the change.

If there are no tracked changes to be found, the current selection remains unchanged.


## Example

This example rejects the next tracked change found after the fifth paragraph in the active document. The  `revTemp`variable is set to  **Nothing** if a change is not found.


```vb
Dim rngTemp as Range 
Dim revTemp as Revision 
 
If ActiveDocument.Paragraphs.Count >= 5 Then 
 Set rngTemp = ActiveDocument.Paragraphs(5).Range 
 rngTemp.Select 
 Set revTemp = Selection.NextRevision(Wrap:=False) 
 If Not (revTemp Is Nothing) Then revTemp.Reject 
End If
```

This example accepts the next tracked change found if the change type is inserted text.




```vb
Dim revTemp as Revision 
 
Set revTemp = Selection.NextRevision(Wrap:=True) 
If Not (revTemp Is Nothing) Then 
 If revTemp.Type = wdRevisionInsert Then revTemp.Accept 
End If
```

This example finds the next revision after the current selection made by the author of the document.




```vb
Dim revTemp as Revision 
Dim strAuthor as String 
 
strAuthor = ActiveDocument.BuiltInDocumentProperties(wdPropertyAuthor) 
 
Do While True 
 Set revTemp = Selection.NextRevision(Wrap:=False) 
 If Not (revTemp Is Nothing) Then 
 If revTemp.Author = strAuthor Then 
 MsgBox Prompt:="Another revision by " &; strAuthor &; "!" 
 Exit Do 
 End If 
 Else 
 MsgBox Prompt:="No more revisions!" 
 Exit Do 
 End If 
Loop
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

