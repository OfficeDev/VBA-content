---
title: Table.TableFields Property (Project)
keywords: vbapj.chm132698
f1_keywords:
- vbapj.chm132698
ms.prod: project-server
api_name:
- Project.Table.TableFields
ms.assetid: 2db4b5fd-6238-b4ab-dc9f-5de991eaad8e
ms.date: 06/08/2017
---


# Table.TableFields Property (Project)

Gets a  **[TableFields](tablefield-object-project.md)** collection representing the fields in the table. Read-only **TableFields**.


## Syntax

 _expression_. **TableFields**

 _expression_ A variable that represents a **Table** object.


## Example

The following example changes the alignment of a column in an entry table. The macro asks for input from the user to indicate which column the user wants to center, then changes the display and refreshes the view.


```vb
Sub AutoWrap() 
 Dim fieldNumber As Integer 
 
 fieldNumber = InputBox$(Prompt:="Enter the number of the " _ 
 &; "column you want to center in the Entry table." _ 
 &; Chr(13) &; "For example, Column 1 is the Indicators " _ 
 &; "column.") 
 
 ActiveProject.TaskTables("Entry").TableFields(fieldNumber _ 
 + 1).AlignData = pjCenter 
 
 TableApply Name:="&;Entry" 
End Sub
```


