---
title: Form.SubdatasheetExpanded Property (Access)
keywords: vbaac10.chm13511
f1_keywords:
- vbaac10.chm13511
ms.prod: access
api_name:
- Access.Form.SubdatasheetExpanded
ms.assetid: 543f2398-ca70-5261-0f9f-e1d864c442e0
ms.date: 06/08/2017
---


# Form.SubdatasheetExpanded Property (Access)

You can use the  **SubdatasheetExpanded** property to specify or determine the saved state of all subdatasheets within a table or query. Read/write **Boolean**.


## Syntax

 _expression_. **SubdatasheetExpanded**

 _expression_ A variable that represents a **Form** object.


## Remarks

The default value is  **False**.

To set the  **SubdatasheetExpanded** property by using Visual Basic, you must first create the property by using the DAO **CreateProperty** method.

The  **SubdatasheetExpanded** and **SubdatasheetHeight** properties take effect on the subform control when the form is in datasheet view.


## Example

The following example turns subdatasheet expansion on or off for the "Purchase Orders" form. 

To try this example yourself, open a form (containing a subform) in Design view, click the Builder button next to the On Load property box in the form's property window, paste this code into the form's Form_Load event (removing the reference to the "Purchase Orders" form), and then open the form in Datasheet view.




```vb
Dim strExpand As String 
 
With Forms("Purchase Orders") 
 
 strExpand = InputBox("Expand subdatasheets? Y/N") 
 
 Select Case strExpand 
 Case "Y" 
 .SubdatasheetExpanded = True 
 Case "N" 
 .SubdatasheetExpanded = False 
 Case Else 
 MsgBox "Can't determine subdatasheet expansion state." 
 End Select 
 
End With
```


## See also


#### Concepts


[Form Object](form-object-access.md)

