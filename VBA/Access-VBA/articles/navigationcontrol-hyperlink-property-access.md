---
title: NavigationControl.Hyperlink Property (Access)
keywords: vbaac10.chm11037
f1_keywords:
- vbaac10.chm11037
ms.prod: access
api_name:
- Access.NavigationControl.Hyperlink
ms.assetid: 2d6bdb1a-808e-1712-1846-71ffa8619f0d
ms.date: 06/08/2017
---


# NavigationControl.Hyperlink Property (Access)

You can use the  **Hyperlink** property to return a reference to a **Hyperlink** object. You can use the **Hyperlink** property to access the properties and methods of a control's hyperlink. Read-only.


## Syntax

 _expression_. **Hyperlink**

 _expression_ A variable that represents a **NavigationControl** object.


## Example

The CreateHyperlink procedure in the following example sets the hyperlink properties for a command button, label, or image control to the address and subaddress values passed to the procedure. The address setting is an optional argument, because a hyperlink to an object in the current database uses only the subaddress setting, To try this example, create a form with two text box controls ( `txtAddress` and `txtSubAddress`) and a command button ( `cmdFollowLink`) and paste the following into the Declarations section of the form's module:


```vb
Private Sub cmdFollowLink_Click() 
 CreateHyperlink Me!cmdFollowLink, Me!txtSubAddress, _ 
 Me!txtAddress 
End Sub 
 
Sub CreateHyperlink(ctlSelected As Control, _ 
 strSubAddress As String, Optional strAddress As String) 
 Dim hlk As Hyperlink 
 Select Case ctlSelected.ControlType 
 Case acLabel, acImage, acCommandButton 
 Set hlk = ctlSelected.Hyperlink 
 With hlk 
 If Not IsMissing(strAddress) Then 
 .Address = strAddress 
 Else 
 .Address = "" 
 End If 
 .SubAddress = strSubAddress 
 .Follow 
 .Address = "" 
 .SubAddress = "" 
 End With 
 Case Else 
 MsgBox "The control '" &; ctlSelected.Name _ 
 &; "' does not support hyperlinks." 
 End Select 
End Sub
```


## See also


#### Concepts


[NavigationControl Object](navigationcontrol-object-access.md)

