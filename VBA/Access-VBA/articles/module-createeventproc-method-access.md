---
title: Module.CreateEventProc Method (Access)
keywords: vbaac10.chm12285
f1_keywords:
- vbaac10.chm12285
ms.prod: access
api_name:
- Access.Module.CreateEventProc
ms.assetid: 13d2a4db-ec80-4225-f3fd-87527dbf660e
ms.date: 06/08/2017
---


# Module.CreateEventProc Method (Access)

The  **CreateEventProc** method creates an event procedure in a class module.


## Syntax

 _expression_. **CreateEventProc**( ** _EventName_**, ** _ObjectName_** )

 _expression_ A variable that represents a **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EventName_|Required|**String**|The name of an event.|
| _ObjectName_|Required|**String**|An object that has the event specified by the  _eventname_ argument. If the event procedure is being added to a **[Form](form-object-access.md)**, the word "Form" should be specified for this argument. If the event procedure is being added to a **[Report](report-object-access.md)**, the word "Report" should be specified for this argument. If the event procedure is being added to a **[Control](control-object-access.md)**, the name of the control should be specified for this argument.|

### Return Value

Long


## Remarks

The value returned by the  **CreateEventProc** method indicates the line number of the first line of the event procedure.

The  **CreateEventProc** method creates a code stub for an event procedure for the specified object. For example, you can use this method to create a Click event procedure for a command button on a form. Microsoft Access creates the Click event procedure in the module associated with the form that contains the command button.

Once you've created the event procedure code stub by using the  **CreateEventProc** method, you can add lines of code to the procedure by using other methods of the **Module** object. For example, you can use the **[InsertLines](module-insertlines-method-access.md)** method to insert a line of code.


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

