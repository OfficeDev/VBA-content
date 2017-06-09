---
title: Module.InsertLines Method (Access)
keywords: vbaac10.chm12277
f1_keywords:
- vbaac10.chm12277
ms.prod: access
api_name:
- Access.Module.InsertLines
ms.assetid: 54ea5ce3-fb2a-e9c7-85ef-8861141f63ec
ms.date: 06/08/2017
---


# Module.InsertLines Method (Access)

The  **InsertLines** method inserts a line or group of lines of code in a standard module or a class module.


## Syntax

 _expression_. **InsertLines**( ** _Line_**, ** _String_** )

 _expression_ A variable that represents a **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Line_|Required|**Long**|The number of the line at which to begin inserting.|
| _String_|Required|**String**|The text to be inserted into the module.|

### Return Value

Nothing


## Remarks

To add multiple lines, include the intrinsic constant **vbCrLf** at the desired line breaks within the string that makes up the _string_ argument. This constant forces a carriage return and line feed.

When you use the  **InsertLines** method, any existing code at the line specified by the _line_ argument moves down.

Lines in a module are numbered beginning with one. To determine the number of lines in a module, use the  **[CountOfLines](module-countoflines-property-access.md)** property.


## Example

The following example creates a new form, adds a command button, and creates a Click event procedure for the command button:


```vb
Function ClickEventProc() As Boolean 
 Dim frm As Form, ctl As Control, mdl As Module 
 Dim lngReturn As Long 
 
 On Error GoTo Error_ClickEventProc 
 ' Create new form. 
 Set frm = CreateForm 
 ' Create command button on form. 
 Set ctl = CreateControl(frm.Name, acCommandButton, , , , _ 
 1000, 1000) 
 ctl.Caption = "Click here" 
 ' Return reference to form module. 
 Set mdl = frm.Module 
 ' Add event procedure. 
 lngReturn = mdl.CreateEventProc("Click", ctl.Name) 
 ' Insert text into body of procedure. 
 mdl.InsertLines lngReturn + 1, vbTab &; "MsgBox ""Way cool!""" 
 ClickEventProc = True 
 
Exit_ClickEventProc: 
 Exit Function 
 
Error_ClickEventProc: 
 MsgBox Err &; " :" &; Err.Description 
 ClickEventProc = False 
 Resume Exit_ClickEventProc 
End Function
```


## See also


#### Concepts


[Module Object](module-object-access.md)

