---
title: ComboBox.ListRows Property (Access)
keywords: vbaac10.chm11384,vbaac10.chm4417
f1_keywords:
- vbaac10.chm11384,vbaac10.chm4417
ms.prod: access
api_name:
- Access.ComboBox.ListRows
ms.assetid: b418e124-71b6-2ffb-101d-b56aadebb1fc
ms.date: 06/08/2017
---


# ComboBox.ListRows Property (Access)

You can use the  **ListRows** property to set the maximum number of rows to display in the list box portion of a combo box. Read/write **Integer**.


## Syntax

 _expression_. **ListRows**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

The  **ListRows** property holds an integer that indicates the maximum number of rows to display. The default setting is 16. The setting for the **ListRows** property must be from 1 to 255.


 **Note**  Microsoft Access sets the  **ListRows** property automatically when you select Lookup Wizard as the data type for a field in table Design view.

You can set the default for this property by using a combo box's default control style or the  **DefaultControl** property in Visual Basic.

If the actual number of rows exceeds the number specified by the  **ListRows** property setting, a vertical scroll bar appears in the list box portion of the combo box.


## Example

The following example uses the  **ListCount** property to find the number of rows in the list box portion of the CustomerList combo box on a Customers form. It then sets the **ListRows** property to display a specified number of rows in the list.


```vb
Public Sub SizeCustomerList() 
 
 Dim ListControl As Control 
 
 Set ListControl = Forms!Customers!CustomerList 
 With ListControl 
 If .ListCount < 8 Then 
 .ListRows = .ListCount 
 Else 
 .ListRows = 8 
 End If 
 End With 
 
End Sub
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

