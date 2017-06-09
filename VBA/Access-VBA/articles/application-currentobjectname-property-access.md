---
title: Application.CurrentObjectName Property (Access)
keywords: vbaac10.chm12502
f1_keywords:
- vbaac10.chm12502
ms.prod: access
api_name:
- Access.Application.CurrentObjectName
ms.assetid: 85b32556-96ed-ed3c-dc5b-4c2570639f50
ms.date: 06/08/2017
---


# Application.CurrentObjectName Property (Access)

You can use the  **CurrentObjectName** property with the **[Application](application-object-access.md)** object to determine the name of the active database object. The active database object is the object that has the focus or in which code is running. Read-only **String**.


## Syntax

 _expression_. **CurrentObjectName**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **CurrentObjectName** property is set by Microsoft Access to a string expression containing the name of the active object.

The following conditions determine which object is considered the active object:


- If the active object is a property sheet, command bar, menu, palette, or field list of an object, the  **CurrentObjectName** property returns the name of the underlying object.
    
- If the active object is a pop-up form, the  **CurrentObjectName** property refers to the pop-up form itself, not the form from which it was opened.
    
- If the active object is the Database window, the  **CurrentObjectName** property returns the item selected in the Database window.
    
- If no object is selected, the  **CurrentObjectName** property returns a zero-length string (" ").
    
- If the current state is ambiguous (the active object isn't a table, query, form, report, macro, or module) for example, if a dialog box has the focus the  **CurrentObjectName** property returns the dialog box name.
    

## Example

You can use this property with the  **[SysCmd](application-syscmd-method-access.md)** method to determine the active object and its state (for example, if the object is open, new, or has been changed but not saved).

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

