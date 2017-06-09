---
title: Rectangle.ControlType Property (Access)
keywords: vbaac10.chm10280
f1_keywords:
- vbaac10.chm10280
ms.prod: access
api_name:
- Access.Rectangle.ControlType
ms.assetid: 08fff4a9-4f15-c65f-5c3e-74d4ef4cf400
ms.date: 06/08/2017
---


# Rectangle.ControlType Property (Access)

You can use the  **ControlType** property in Visual Basic to determine the type of a control on a form or report. Read/write **Byte**.


## Syntax

 _expression_. **ControlType**

 _expression_ A variable that represents a **Rectangle** object.


## Remarks

The  **ControlType** property setting is an intrinsic constant that specifies the control type.



|**Constant**|**Control**|
|:-----|:-----|
|**acBoundObjectFrame**|[Bound object frame](bound-object-frame-control.md)|
|**acCheckBox**|[Check box](check-box-control.md)|
|**acComboBox**|[Combo box](combo-box-control.md)|
|**acCommandButton**|[Command button](command-button-control.md)|
|**acCustomControl**|[ActiveX (custom) control](activex-control.md)|
|**acImage**|[Image](image-control-misc.md)|
|**acLabel**|[Label](label-control-access.md)|
|**acLine**|[Line](line-control.md)|
|**acListBox**|[List box](list-box-control.md)|
|**acObjectFrame**|[Unbound object frame](unbound-object-frame-control.md)or [Chart](chart-control.md)|
|**acOptionButton**|[Option button](option-button-control.md)|
|**acOptionGroup**|[Option group](option-group-control.md)|
|**acPage**|[Page](page.md)|
|**acPageBreak**|[Page break](page-break-control.md)|
|**acRectangle**|[Rectangle](rectangle-control.md)|
|**acSubform**|[Subform/subreport](subform-subreport-control.md)|
|**acTabCtl**|[Tab](tab-control.md)|
|**acTextBox**|[Text box](text-box-control.md)|
|**acToggleButton**|[Toggle button](toggle-button-control.md)[Toggle button](toggle-button-control.md)|

 **Note**  The  **ControlType** property can only be set by using Visual Basic in form Design view or report Design view, but it can be read in all views.

The  **ControlType** property is useful not only for checking for a specific control type in code, but also for changing the type of control to another type. For example, you can change a text box to a combo box by setting the **ControlType** property for the text box to **acComboBox** while in form Design view.

You can use the  **ControlType** property to change characteristics of similar controls on a form according to certain conditions. For example, if you don't want users to edit existing data in text boxes, you can set the **SpecialEffect** property for all text boxes to Flat and set the form's **AllowEdits** property to No. (The **SpecialEffect** property doesn't affect whether data can be edited; it's used here to provide a visual cue that the control behavior has changed.)

The  **ControlType** property is also used to specify the type of control to create when you are using the **[CreateControl](application-createcontrol-method-access.md)** method.


## Example

The following example examines the  **ControlType** property for all controls on a form. For each label and text box control, the procedure toggles the **SpecialEffect** property for those controls. When the label controls' **SpecialEffect** property is set to Shadowed and the text box controls' **SpecialEffect** property is set to Normal and the **AllowAdditions**, **AllowDeletions**, and **AllowEdits** properties are all set to **True**, the `intCanEdit` variable is toggled to allow editing of the underlying data.


```vb
Sub ToggleControl(frm As Form) 
 Dim ctl As Control 
 Dim intI As Integer, intCanEdit As Integer 
 Const conTransparent = 0 
 Const conWhite = 16777215 
 For Each ctl in frm.Controls 
 With ctl 
 Select Case .ControlType 
 Case acLabel 
 If .SpecialEffect = acEffectShadow Then 
 .SpecialEffect = acEffectNormal 
 .BorderStyle = conTransparent 
 intCanEdit = True 
 Else 
 .SpecialEffect = acEffectShadow 
 intCanEdit = False 
 End If 
 Case acTextBox 
 If .SpecialEffect = acEffectNormal Then 
 .SpecialEffect = acEffectSunken 
 .BackColor = conWhite 
 Else 
 .SpecialEffect = acEffectNormal 
 .BackColor = frm.Detail.BackColor 
 End If 
 End Select 
 End With 
 Next ctl 
 If intCanEdit = IFalse Then 
 With frm 
 .AllowAdditions = False 
 .AllowDeletions = False 
 .AllowEdits = False 
 End With 
 Else 
 With frm 
 .AllowAdditions = True 
 .AllowDeletions = True 
 .AllowEdits = True 
 End With 
 End If 
End Sub
```


## See also


#### Concepts


[Rectangle Object](rectangle-object-access.md)

