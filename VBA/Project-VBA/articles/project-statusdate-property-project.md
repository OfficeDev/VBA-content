---
title: Project.StatusDate Property (Project)
keywords: vbapj.chm132611
f1_keywords:
- vbapj.chm132611
ms.prod: project-server
api_name:
- Project.Project.StatusDate
ms.assetid: 3d53790c-051c-e3d1-887a-1329c8ef98a8
ms.date: 06/08/2017
---


# Project.StatusDate Property (Project)

Gets or sets the current status date for the project. If there is no status date, returns "NA". Read/write  **Variant**.


## Syntax

 _expression_. **StatusDate**

 _expression_ A variable that represents a **Project** object.


## Example

The following example checks the status date of the active project. If it has never changed from the default, but edits to total task percent complete are spread to the status date, it asks for a status date to use. If edits to total task percent complete are spread to the calculated stop date of the task, it asks the user if the edits should be spread to a status date instead and, if so, asks for a status date to use.


```vb
Sub SpreadPercentComplete() 
 Dim NewStatus As Date, AskToSpread As Long 
 
 With ActiveProject 
 If .StatusDate = "NA" And .SpreadPercentCompleteToStatusDate Then 
 NewStatus = InputBox("Enter a status date for the project: ") 
 .StatusDate = NewStatus 
 MsgBox "The status date was set to " &; .StatusDate &; "." 
 ElseIf .SpreadPercentCompleteToStatusDate = False Then 
 AskToSpread = MsgBox("Should changes to total task percent complete" &; _ 
 " be spread to a status date?", vbYesNo) 
 If AskToSpread = vbYes Then 
 NewStatus = InputBox("Enter a status date for the project: ") 
 .StatusDate = NewStatus 
 .SpreadPercentCompleteToStatusDate = True 
 MsgBox "The status date was set to " &; .StatusDate &; "." 
 End If 
 End If 
 End With 
 
End Sub
```


