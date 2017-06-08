---
title: Form.MinMaxButtons Property (Access)
keywords: vbaac10.chm13375
f1_keywords:
- vbaac10.chm13375
ms.prod: access
api_name:
- Access.Form.MinMaxButtons
ms.assetid: 12f2a0b1-1f45-544b-b116-8d5aa51d6897
ms.date: 06/08/2017
---


# Form.MinMaxButtons Property (Access)

You can use the  **MinMaxButtons** property to specify whether the **Maximize** and **Minimize** buttons will be visible on a form. Read/write **Byte**.


## Syntax

 _expression_. **MinMaxButtons**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **MinMaxButtons** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|None|0|The  **Maximize** and **Minimize** buttons aren't visible.|
|Min Enabled|1|Only the  **Minimize** button is visible.|
|Max Enabled|2|Only the  **Maximize** button is visible.|
|Both Enabled|3|(Default) Both the  **Minimize** and **Maximize** buttons are visible.|
You can set the  **MinMaxButtons** property only in form Design view.

Clicking a form's  **Maximize** button enlarges the form so it fills the Microsoft Access window. Clicking a form's **Minimize** button reduces the form to a short title bar at the bottom of the Microsoft Access window.

To display the  **Maximize** and **Minimize** buttons on a form, you must set the form's **BorderStyle** property to Thin or Sizable and the **ControlBox** property to Yes. If you set the **BorderStyle** property to None or Dialog, or if you set the **ControlBox** property to No, the form won't have **Maximize** or **Minimize** buttons, regardless of the **MinMaxButtons** property setting.

Even when the  **MinMaxButtons** property is set to None, a form always has **Maximize** and **Minimize** buttons in Design view.

If a form's  **MinMaxButtons** property is set to None, the **Maximize** and **Minimize** commands aren't available on the form's **Control** menu.


## Example

The following example returns the value of the  **MinMaxButtons** property for the "Order Entry" form.


```vb
Dim b As Byte 
b = Forms("Order Entry").MinMaxButtons
```


## See also


#### Concepts


[Form Object](form-object-access.md)

