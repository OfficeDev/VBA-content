---
title: Permission.RemoveAll Method (Office)
keywords: vbaof11.chm261007
f1_keywords:
- vbaof11.chm261007
ms.prod: office
api_name:
- Office.Permission.RemoveAll
ms.assetid: 33dc3f62-c92f-03b0-e164-98c366bbdb32
ms.date: 06/08/2017
---


# Permission.RemoveAll Method (Office)

Removes all  **UserPermission** objects from the **Permission** collection of the active document.


## Syntax

 _expression_. **RemoveAll**

 _expression_ A variable that represents a **Permission** object.


## Remarks

The  **RemoveAll** method removes all UserPermissions that have been added to the **Permission** collection and disables restrictions on the active document. After calling the **RemoveAll** method, the **Enabled** property of the **Permission** object returns **False** and the **Count** property returns 0 (zero).


## Example

The following example uses the  **RemoveAll** method to remove all user permissions and to disable restrictions on the active document.


```
 Dim irmPermission As Office.Permission 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 irmPermission.RemoveAll 
 MsgBox "All permissions removed." &amp; vbCrLf &amp; _ 
 "Count: " &amp; irmPermission.Count &amp; vbCrLf &amp; _ 
 "Enabled: " &amp; irmPermission.Enabled, _ 
 vbInformation + vbOKOnly, "IRM Information" 
 Else 
 MsgBox "This document is not restricted.", _ 
 vbInformation + vbOKOnly, "IRM Information" 
 End If 
 Set irmPermission = Nothing 

```


## See also


#### Concepts


[Permission Object](permission-object-office.md)
#### Other resources


[Permission Object Members](permission-members-office.md)

