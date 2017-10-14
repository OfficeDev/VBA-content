---
title: Conflict.Reject Method (Word)
keywords: vbawd10.chm78708838
f1_keywords:
- vbawd10.chm78708838
ms.prod: word
api_name:
- Word.Conflict.Reject
ms.assetid: 9bd4fa93-4bae-e2a8-ef6e-b3116542cad4
ms.date: 06/08/2017
---


# Conflict.Reject Method (Word)

Rejects the user change, removes the conflict, and accepts the server copy of the change for the conflict.


## Syntax

 _expression_ . **Reject**

 _expression_ An expression that returns a **Conflict** object.


### Return Value

Nothing


## Remarks

The  **Reject** method rejects the user version of a conflict and accepts the version that is currently on the server.


## Example

The following code example rejects all the conflicts in the active document.


```vb
Dim conf As Conflict 
 
For Each conf In ActiveDocument.CoAuthoring.Conflicts 
 conf.Reject 
Next conf
```

Alternatively, you can use the [RejectAll](conflicts-rejectall-method-word.md) method of the[Conflicts](conflicts-object-word.md) collection object to reject all the conflicts in a document, as shown in the following code example.




```vb
ActiveDocument.CoAuthoring.Conflicts.RejectAll
```


## See also


#### Concepts


[Conflict Object](conflict-object-word.md)

