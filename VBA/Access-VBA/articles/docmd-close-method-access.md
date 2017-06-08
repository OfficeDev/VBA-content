---
title: DoCmd.Close Method (Access)
keywords: vbaac10.chm4145
f1_keywords:
- vbaac10.chm4145
ms.prod: access
api_name:
- Access.DoCmd.Close
ms.assetid: 3fdb2fa2-31d8-baf7-89f3-f9ef330280b3
ms.date: 06/08/2017
---


# DoCmd.Close Method (Access)

The  **Close** method carries out the Close action in Visual Basic.


## Syntax

 _expression_. **Close**( ** _ObjectType_**, ** _ObjectName_**, ** _Save_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Optional|**AcObjectType**|A  **[AcObjectType](acobjecttype-enumeration-access.md)** constant that represents the type of object to close.|
| _ObjectName_|Optional|**Variant**|A string expression that's the valid name of an object of the type selected by the  _objecttype_ argument.|
| _Save_|Optional|**AcCloseSave**|A  **[AcCloseSave](acclosesave-enumeration-access.md)** constant tha specifies whether or not to save changes to the object. The default value is **acSavePrompt**.|

## Remarks

You can use the  **Close** method to close either a specified Microsoft Access window or the active window if none is specified.

If you leave the  _objecttype_ and _objectname_ arguments blank (the default constant, **acDefault**, is assumed for _objecttype_), Microsoft Access closes the active window. If you specify the  _save_ argument and leave the _objecttype_ and _objectname_ arguments blank, you must include the _objecttype_ and _objectname_ arguments' commas.


 **Note**   If a form has a control bound to a field that has its **Required** property set to 'Yes,' and the form is closed using the **Close** method without entering any data for that field, an error message is not displayed. Any changes made to the record will be aborted. When the form is closed using the user interface, Microsoft Access displays an alert.

To display an error message, use the  **RunCommand** method to invoke the **acCmdSaveRecord** command before calling the **Close** method. This will cause a run-time error if one or more required fields are Null. This technique is illustrated in the following example.




```vb
Private Sub cmdCloseForm_Click() 
On Error GoTo Err_cmdCloseForm_Click 
 
 DoCmd.RunCommand acCmdSaveRecord 
 DoCmd.Close 
 
Exit_cmdCloseForm_Click: 
 Exit Sub 
 
Err_cmdCloseForm_Click: 
 MsgBox Err.Description 
 Resume Exit_cmdCloseForm_Click 
 
End Sub
```


## Example

The following example uses the  **Close** method to close the form Order Review, saving any changes to the form without prompting:


```vb
DoCmd.Close acForm, "Order Review", acSaveYes
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

