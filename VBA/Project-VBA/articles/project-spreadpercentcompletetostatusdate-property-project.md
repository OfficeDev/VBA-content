---
title: Project.SpreadPercentCompleteToStatusDate Property (Project)
ms.prod: project-server
api_name:
- Project.Project.SpreadPercentCompleteToStatusDate
ms.assetid: c1c9a8eb-8572-7bad-33b2-23157c908f60
ms.date: 06/08/2017
---


# Project.SpreadPercentCompleteToStatusDate Property (Project)

 **True** if edits to total task percent complete are spread to the status date, or to the current date if the status date is "NA". **False** if edits are spread to the calculated stop date of the task. Read/write **Boolean**.


## Syntax

 _expression_. **SpreadPercentCompleteToStatusDate**

 _expression_ A variable that represents a **Project** object.


## Example

The following example checks the status date of the active project. If it has never changed from the default, but edits to total task percent complete are spread to the status date, the macro asks for a status date to use. If edits to total task percent complete are spread to the calculated stop date of the task, the macro asks the user if edits should be spread to a status date instead and, if so, asks for a status date to use.


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


