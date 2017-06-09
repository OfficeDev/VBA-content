---
title: Application.CreateControl Method (Access)
keywords: vbaac10.chm12622
f1_keywords:
- vbaac10.chm12622
ms.prod: access
api_name:
- Access.Application.CreateControl
ms.assetid: f5b1689c-62c4-163d-c659-607cee7572f6
ms.date: 06/08/2017
---


# Application.CreateControl Method (Access)

The  **CreateControl** method creates a control on a specified open form. For example, suppose you are building a custom wizard that allows users to easily construct a particular form. You can use the **CreateControl** method in your wizard to add the appropriate controls to the form.


## Syntax

 _expression_. **CreateControl**( ** _FormName_**, ** _ControlType_**, ** _Section_**, ** _Parent_**, ** _ColumnName_**, ** _Left_**, ** _Top_**, ** _Width_**, ** _Height_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FormName_|Required|**String**|The name of the open form or report on which you want to create the control.|
| _ControlType_|Required|**AcControlType**|An  **[AcControlType](accontroltype-enumeration-access.md)** constant that represents the type of control you want to create.|
| _Section_|Optional|**AcSection**|An  **[AcSection](acsection-enumeration-access.md)** constant that identifying the section that will contain the new control.|
| _Parent_|Optional|**Variant**|The name of the parent control of an attached control. For controls that have no parent control, use a zero-length string for this argument, or omit it.|
| _ColumnName_|Optional|**Variant**|The name of the field to which the control will be bound, if it is to be a data-bound control.|
| _Left,Top_|Optional|**Variant**|The coordinates for the upper-left corner of the control in twips.|
| _Width, Height_|Optional|**Variant**|Numeric expressions indicating the width and height of the control in twips.|

### Return Value

Control


## Remarks

You can use the Create **Control** and **CreateReportControl** methods in a custom wizard to create controls on a form or report. Both methods return a **[Control](control-object-access.md)** object.

You can use the  **CreateControl** and **CreateReportControl** methods only in form Design view or report Design view, respectively.

You use the  _parent_ argument to identify the relationship between a main control and a subordinate control. For example, if a text box has an attached label, the text box is the main (or parent) control and the label is the subordinate (or child) control. When you create the label control, set its _parent_ argument to a string identifying the name of the parent control. When you create the text box, set its _parent_ argument to a zero-length string.

You also set the  _parent_ argument when you create check boxes, option buttons, or toggle buttons. An option group is the parent control of any check boxes, option buttons, or toggle buttons that it contains. The only controls that can have a parent control are a label, check box, option button, or toggle button. All of these controls can also be created independently, without a parent control.

Set the  _columnname_ argument according to the type of control you are creating and whether or not it will be bound to a field in a table. The controls that may be bound to a field include the text box, list box, combo box, option group, and bound object frame. Additionally, the toggle button, option button, and check box controls may be bound to a field if they are not contained in an option group.

If you specify the name of a field for the  _columnname_ argument, you create a control that is bound to that field. All of the control's properties are then automatically set to the settings of any corresponding field properties. For example, the value of the control's **ValidationRule** property will be the same as the value of that property for the field.


 **Note**  If your wizard creates controls on a new or existing form or report, it must first open the form or report in Design view.

To remove a control from a form or report, use the  **[DeleteControl](application-deletecontrol-method-access.md)** and **[DeleteReportControl](application-deletereportcontrol-method-access.md)** statements.


## Example

The following example first creates a new form based on an Orders table. It then uses the  **CreateControl** method to create a text box control and an attached label control on the form.


```vb
Sub NewControls() 
 Dim frm As Form 
 Dim ctlLabel As Control, ctlText As Control 
 Dim intDataX As Integer, intDataY As Integer 
 Dim intLabelX As Integer, intLabelY As Integer 
 
 ' Create new form with Orders table as its record source. 
 Set frm = CreateForm 
 frm.RecordSource = "Orders" 
 ' Set positioning values for new controls. 
 intLabelX = 100 
 intLabelY = 100 
 intDataX = 1000 
 intDataY = 100 
 ' Create unbound default-size text box in detail section. 
 Set ctlText = CreateControl(frm.Name, acTextBox, , "", "", _ 
 intDataX, intDataY) 
 ' Create child label control for text box. 
 Set ctlLabel = CreateControl(frm.Name, acLabel, , _ 
 ctlText.Name, "NewLabel", intLabelX, intLabelY) 
 ' Restore form. 
 DoCmd.Restore 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

