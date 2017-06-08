---
title: BoundObjectFrame.StatusBarText Property (Access)
keywords: vbaac10.chm10920
f1_keywords:
- vbaac10.chm10920
ms.prod: access
api_name:
- Access.BoundObjectFrame.StatusBarText
ms.assetid: dc8ec458-8013-f6ff-5763-d083babcb4c9
ms.date: 06/08/2017
---


# BoundObjectFrame.StatusBarText Property (Access)

You can use the  **StatusBarText** property to specify the text that is displayed in the status bar when a control is selected. Read/write **String**.


## Syntax

 _expression_. **StatusBarText**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

You set the  **StatusBarText** property by using a string expression up to 255 characters long.


 **Note**  The length of the text you can display in the status bar depends on your computer hardware and video display.

You can use the  **StatusBarText** property to provide specific information about a control. For example, when a text box has the focus, a brief instruction can tell the user what kind of data to enter


 **Note**  You can also use the  **ControlTipText** property to display a ScreenTip for a control.

If you create a control by dragging a field from the field list, the value in a field's  **Description** property is copied to the **StatusBarText** property.


## Example

The following example sets the status bar text to be displayed when the "Address_TextBox" control in the "Mailing List" form has the focus in Form View. 


```vb
Forms("Mailing List").Controls("Address_TextBox"). _ 
 StatusBarText = "Enter the company's mailing address." 

```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

