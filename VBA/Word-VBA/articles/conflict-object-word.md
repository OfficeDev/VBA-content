---
title: Conflict Object (Word)
keywords: vbawd10.chm1201
f1_keywords:
- vbawd10.chm1201
ms.prod: word
api_name:
- Word.Conflict
ms.assetid: e9fe0318-d3e3-7589-0c15-64210ac5b709
ms.date: 06/08/2017
---


# Conflict Object (Word)

Represents a conflicting edit in a co authored document. The type of a  **Conflict** object is specified by the[WdRevisionType](wdrevisiontype-enumeration-word.md) enumeration.


## Remarks

Although co authoring in Word is designed to minimize conflicts, conflicts can sometimes occur when editing a document that has co authoring enabled. A conflict occurs when Word requires user input to resolve a merge.


 **Note**  Documents can only be co authored on a server that supports the File Synchronization via SOAP over HTTP protocol, such as Microsoft SharePoint Server 2010.

For example, conflicts could potentially occur when a user opens a co authored document from the server, works offline, and once online again, saves the document back to the server. As another example, conflicts can sometimes occur when more than one person works on the same document range at exactly the same time.


 **Note**  A user is only made aware of conflicts in the document when they perform an explicit document save. When the user performs an explicit document save, Word will enter Conflict Resolution mode if there are conflicts in the document. Conflict Resolution mode enables the user to resolve document conflicts. 


## Example

The following code example gets the type of each conflict in the active document.


```vb
Dim con as Conflict 
 
For Each con in ActiveDocument.CoAuthoring.Conflicts 
MsgBox con.Type 
Next con
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

