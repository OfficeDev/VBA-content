---
title: MailMergeFields Object (Word)
ms.prod: word
ms.assetid: 9d2dfd45-c52b-500e-15bf-1e678e6c1e92
ms.date: 06/08/2017
---


# MailMergeFields Object (Word)

A collection of  **[MailMergeField](mailmergefield-object-word.md)** objects that represent the mail merge related fields in a document.


## Remarks

Use the  **Fields** property to return the **MailMergeFields** collection. The following example adds an ASK field after the last mail merge field in the active document.


```
Set myMMFields = ActiveDocument.MailMerge.Fields 
myMMFields(myMMFields.Count).Select 
Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdMove 
ActiveDocument.MailMerge.Fields.AddAsk Range:=Selection.Range, _ 
 Name:="Name", Prompt:="Type your name", AskOnce:=True
```

Use the  **Add** method to add a merge field to the **MailMergeFields** collection. The following example replaces the selection with a **MiddleInitial** merge field.




```
ActiveDocument.MailMerge.Fields.Add Range:=Selection.Range, _ 
 Name:="MiddleInitial"
```

Use  **Fields** (Index), where Index is the index number, to return a single **MailMergeField** object. The following example displays the field code of the first mail merge field in the active document.




```
MsgBox ActiveDocument.MailMerge.Fields(1).Code
```

The  **MailMergeFields** collection has additional methods, such as **AddAsk** and **AddFillIn**, for adding fields related to a mail merge operation.


## Methods



|**Name**|
|:-----|
|[Add](mailmergefields-add-method-word.md)|
|[AddAsk](mailmergefields-addask-method-word.md)|
|[AddFillIn](mailmergefields-addfillin-method-word.md)|
|[AddIf](mailmergefields-addif-method-word.md)|
|[AddMergeRec](mailmergefields-addmergerec-method-word.md)|
|[AddMergeSeq](mailmergefields-addmergeseq-method-word.md)|
|[AddNext](mailmergefields-addnext-method-word.md)|
|[AddNextIf](mailmergefields-addnextif-method-word.md)|
|[AddSet](mailmergefields-addset-method-word.md)|
|[AddSkipIf](mailmergefields-addskipif-method-word.md)|
|[Item](mailmergefields-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](mailmergefields-application-property-word.md)|
|[Count](mailmergefields-count-property-word.md)|
|[Creator](mailmergefields-creator-property-word.md)|
|[Parent](mailmergefields-parent-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
