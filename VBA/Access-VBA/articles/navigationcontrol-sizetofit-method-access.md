---
title: NavigationControl.SizeToFit Method (Access)
keywords: vbaac10.chm11033
f1_keywords:
- vbaac10.chm11033
ms.prod: access
api_name:
- Access.NavigationControl.SizeToFit
ms.assetid: 690d17ca-866d-2f8e-fc54-a5cc166b6ad1
ms.date: 06/08/2017
---


# NavigationControl.SizeToFit Method (Access)

You can use the  **SizeToFit** method to size a control so it fits the text or image that it contains.


## Syntax

 _expression_. **SizeToFit**

 _expression_ A variable that represents a **NavigationControl** object.


## Remarks

For example, you can apply the  **SizeToFit** method to a command button that is too small to display all the text in its **Caption** property.

You can apply the  **SizeToFit** method to controls only in form Design view or report Design view.

The  **SizeToFit** method will make a control larger or smaller, depending on the size of the text or image it contains.

You can use the  **SizeToFit** method in conjunction with the **[CreateControl](application-createcontrol-method-access.md)** method to size new controls that you have created in code.




 **Note**  Not all controls that contain text or an image can be sized by the  **SizeToFit** method. Several controls are bound to data that can vary in size from one record to the next. These controls include the text box, list box, combo box, and bound object frame controls. The **SizeToFit** method does not apply to controls on data access pages.


## Example

The following example creates a new form and creates a new command button on the form. The procedure then sets the control's  **Caption** property and sizes the control to fit the caption.


```vb
Sub SizeNewControl() 
 Dim frm As Form, ctl As Control 
 
 ' Create new form. 
 Set frm = CreateForm 
 ' Create new command button. 
 Set ctl = CreateControl(frm.Name, _ 
 acCommandButton, , , , 500, 500) 
 ' Restore form. 
 DoCmd.Restore 
 ' Set control's Caption property. 
 ctl.Caption = "Extremely Long Control Caption" 
 ' Size control to fit caption. 
 ctl.SizeToFit 
End Sub
```


## See also


#### Concepts


[NavigationControl Object](navigationcontrol-object-access.md)

