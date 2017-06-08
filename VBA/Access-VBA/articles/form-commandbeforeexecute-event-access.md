---
title: Form.CommandBeforeExecute Event (Access)
keywords: vbaac10.chm13673
f1_keywords:
- vbaac10.chm13673
ms.prod: access
api_name:
- Access.Form.CommandBeforeExecute
ms.assetid: 4fb1c072-3781-8a52-bc9a-2e26d2738789
ms.date: 06/08/2017
---


# Form.CommandBeforeExecute Event (Access)

Occurs before a specified command is executed. Use this event when you want to impose certain restrictions before a particular command is executed.


## Syntax

 _expression_. **CommandBeforeExecute**( ** _Command_**, ** _Cancel_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Command_|Required|**Variant**| The command that is going to be executed.|
| _Cancel_|Required|**Object**| Set the **Value** property of this object to **True** to cancel the command.|

### Return Value

nothing


## Remarks

The  **OCCommandId**, **ChartCommandIdEnum**, and **PivotCommandId** constants contain lists of the supported commands for each of the Microsoft Office Web Components.


## Example

The following example demonstrates the syntax for a subroutine that traps the  **CommandBeforeExecute** event.


```vb
Private Sub Form_CommandBeforeExecute( _ 
 ByVal Command As Variant, ByVal Cancel As Object) 
 Dim intResponse As Integer 
 Dim strPrompt As String 
 
 strPrompt = "Cancel the command?" 
 
 intResponse = MsgBox(strPrompt, vbYesNo) 
 
 If intResponse = vbYes Then 
 Cancel.Value = True 
 Else 
 Cancel.Value = False 
 End If 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

