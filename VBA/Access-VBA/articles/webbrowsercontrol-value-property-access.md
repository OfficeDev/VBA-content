---
title: WebBrowserControl.Value Property (Access)
keywords: vbaac10.chm14358
f1_keywords:
- vbaac10.chm14358
ms.prod: access
api_name:
- Access.WebBrowserControl.Value
ms.assetid: bf08215c-14c7-b2b2-65d5-707478e96e5a
ms.date: 06/08/2017
---


# WebBrowserControl.Value Property (Access)

Determines or specifies the text in the text box. Read/write  **Variant**.


## Syntax

 _expression_. **Value**

 _expression_ A variable that represents a **WebBrowserControl** object.


## Remarks

The  **Text** property returns the formatted string. The **Text** property may be different than the **Value** property for a text box control. The **Text** property is the current contents of the control. The **Value** property is the saved value of the text box control. The **Text** property is always current while the control has the focus.

The  **Value** property returns or sets a control's default property, which is the property that is assumed when you don't explicitly specify a property name.


 **Note**   The **Value** property is not the same as the **DefaultValue** property, which specifies the value that a property is assigned when a new record is created.


## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

