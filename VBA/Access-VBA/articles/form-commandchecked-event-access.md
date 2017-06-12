---
title: Form.CommandChecked Event (Access)
keywords: vbaac10.chm13674
f1_keywords:
- vbaac10.chm13674
ms.prod: access
api_name:
- Access.Form.CommandChecked
ms.assetid: ec30f538-bbd2-9935-1ad9-5210f457b15f
ms.date: 06/08/2017
---


# Form.CommandChecked Event (Access)

Occurs when the specified Microsoft Office Web Component determines whether the specified command is checked.


## Syntax

 _expression_. **CommandChecked**( ** _Command_**, ** _Checked_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Command_|Required|**Variant**| The command that has been verified as being checked.|
| _Checked_|Required|**Object**| Set the **Value** property of this object to **False** to uncheck the command.|

### Return Value

nothing


## Remarks

The  **OCCommandId**, **ChartCommandIdEnum**, and **PivotCommandId** constants contain lists of the supported commands for each of the Microsoft Office Web Components.


## Example

The following example demonstrates the syntax for a subroutine that traps the  **CommandChecked** event.


```vb
Private Sub Form_CommandChecked( _ 
 ByVal Command As Variant, ByVal Checked As Object) Dim intResponse As Integer 
 Dim strPrompt As String 
 
 strPrompt = "Uncheck the command?" 
 
 intResponse = MsgBox(strPrompt, vbYesNo) 
 
 If intResponse = vbYes Then 
 Checked.Value = False 
 Else 
 Checked.Value = True 
 End If 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

