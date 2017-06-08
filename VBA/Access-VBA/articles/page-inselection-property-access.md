---
title: Page.InSelection Property (Access)
keywords: vbaac10.chm12166
f1_keywords:
- vbaac10.chm12166
ms.prod: access
api_name:
- Access.Page.InSelection
ms.assetid: 7612601b-da5d-bb0a-4e7e-603b444569be
ms.date: 06/08/2017
---


# Page.InSelection Property (Access)

You can use the  **InSelection** property to determine or specify whether a control on a form in Design view is selected. Read/write **Boolean**.


## Syntax

 _expression_. **InSelection**

 _expression_ A variable that represents a **Page** object.


## Remarks

When a control is selected, its sizing handles are visible and it can be resized by the user. More than one control can be selected at a time.


## Example

The following function uses the  **InSelection** property to determine whether the `strControlName` control on a form is selected.

To test this code, paste the IsControlSelected function code in the Declarations section of a code module in the Northwind sample database, open the Customers form in Design view, and select the CompanyName control. Then enter the following line in the Debug window:




```vb
? IsControlSelected (Forms!Customers, "CompanyName") 
 
Function IsControlSelected(frm As Form, _ 
 strControlName As String) As Boolean 
 Dim intI As Integer, ctl As Control 
 If frm.CurrentView <> 0 Then 
 ' Form is not in Design view. 
 Exit Function 
 Else 
 For intI = 0 To frm.Count - 1 
 Set ctl = frm(intI) 
 If ctl.InSelection = True Then 
 ' Is desired control selected? 
 If UCase(ctl.Name) = UCase(strControlName) Then 
 IsControlSelected = True 
 Exit Function 
 End If 
 Else 
 IsControlSelected = False 
 End If 
 Next intI 
 End If 
End Function
```


## See also


#### Concepts


[Page Object](page-object-access.md)

