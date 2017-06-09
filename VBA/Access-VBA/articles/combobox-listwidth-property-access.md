---
title: ComboBox.ListWidth Property (Access)
keywords: vbaac10.chm11385,vbaac10.chm4418
f1_keywords:
- vbaac10.chm11385,vbaac10.chm4418
ms.prod: access
api_name:
- Access.ComboBox.ListWidth
ms.assetid: 488a36f0-3ab1-1bb1-ff48-3e5d33a55139
ms.date: 06/08/2017
---


# ComboBox.ListWidth Property (Access)

You can use the  **ListWidth** property to set the width of the list box portion of a combo box. Read/write **String**.


## Syntax

 _expression_. **ListWidth**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

The  **ListWidth** property holds a value specifying the width of the list box portion of a combo box in inches or centimeters, depending on the measurement system (U.S. or Metric) selected in the **Measurement system** box on the **Numbers** tab of the **Regional Options** dialog box of Windows Control Panel. To use a unit other than the default, include a measurement indicator, such as cm or in. The default setting (Auto) makes the list box portion of the combo box the same width as the combo box.

The default unit of measurement is twips.


 **Note**  Microsoft Access sets the  **ListWidth** property automatically when you select Lookup Wizard as the data type for a field in table Design view.

You can also set the default for this property by using a combo box's default control style or the  **DefaultControl** property in Visual Basic.

The list portion of the combo box can be wider than the combo box but can't be narrower.

If you want to display a multiple-column list, enter a value that will make the list box wide enough to show all the columns. When designing combo boxes, be sure to leave enough space to display your data and for Microsoft Access to insert a vertical scroll bar.


## Example

The following example returns the value of the  **ListWidth** property for the "States" combo box on the "Order Entry" form.


```vb
Dim str As String 
str = Forms("Order Entry").Controls("States").ListWidth
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

