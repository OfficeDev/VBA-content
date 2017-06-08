---
title: Application.CurrentObjectType Property (Access)
keywords: vbaac10.chm12501
f1_keywords:
- vbaac10.chm12501
ms.prod: access
api_name:
- Access.Application.CurrentObjectType
ms.assetid: 10065578-b218-8b83-f210-056922a57c4b
ms.date: 06/08/2017
---


# Application.CurrentObjectType Property (Access)

You can use the  **CurrentObjectType** property together with the **[Application](application-object-access.md)** object to determine the type of the active database object (table, query, form, report, macro, module, server view, database diagram, or stored procedure). The active database object is the object that has the focus or in which code is running. Read-only **[AcObjectType](acobjecttype-enumeration-access.md)**.


## Syntax

 _expression_. **CurrentObjectType**

 _expression_ A variable that represents an **Application** object.


## Remarks

The following conditions determine which object is considered the active object:


- If the active object is a property sheet, command bar, menu, palette, or field list of an object, the  **CurrentObjectType** property returns the type of the underlying object.
    
- If the active object is a pop-up form, the  **CurrentObjectType** property refers to the pop-up form itself, not the form from which it was opened.
    
- If the active object is the Database window, the  **CurrentObjectType** property returns the item selected in the Database window.
    
- If no object is selected, the  **CurrentObjectType** property returns **True**.
    
- If the current state is ambiguous (the active object isn't a table, query, form, report, macro, or module) for example, if a dialog box has the focus the  **CurrentObjectType** property returns **True**.
    
You can use this property with the  **[SysCmd](application-syscmd-method-access.md)** method to determine the active object and its state (for example, if the object is open, new, or has been changed but not saved).


## Example

The following example uses the  **CurrentObjectType** and **CurrentObjectName** properties with the **SysCmd** function to determine if the active object is the Products form and if this form is open and has been changed but not saved. If these conditions are true, the form is saved and then closed.


```vb
Public Sub CheckProducts() 
 
 Dim intState As Integer 
 Dim intCurrentType As Integer 
 Dim strCurrentName As String 
 
 intCurrentType = Application.CurrentObjectType 
 strCurrentName = Application.CurrentObjectName 
 
 If intCurrentType = acForm And strCurrentName = "Products" Then 
 intState = SysCmd(acSysCmdGetObjectState, intCurrentType, _ 
 strCurrentName) 
 
 ' Products form changed but not saved. 
 If intState = acObjStateDirty + acObjStateOpen Then 
 
 ' Close Products form and save changes. 
 DoCmd.Close intCurrentType, strCurrentName, acSaveYes 
 End If 
 End If 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

