---
title: TabControl.Style Property (Access)
keywords: vbaac10.chm12087
f1_keywords:
- vbaac10.chm12087
ms.prod: access
api_name:
- Access.TabControl.Style
ms.assetid: de0859cd-27af-294b-da0c-ef2055180b21
ms.date: 06/08/2017
---


# TabControl.Style Property (Access)

You can use the  **Style** property to specify or determine the appearance of tabs on a tab control. Read/write **Byte**.


## Syntax

 _expression_. **Style**

 _expression_ A variable that represents a **TabControl** object.


## Remarks

The  **Style** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Tabs|**0**|(Default) Tabs appear as tabs.|
|Buttons|**1**|Tabs appear as buttons.|
|None|**2**|No tabs appear in the control.|
You can also set the default for this property by setting a control's  **DefaultControl** property in Visual Basic.

You can set the  **Style** property in any view.

When the tab control's  **Style** property is set to Tabs or Buttons, the appearance of the tabs is determined by the **TabFixedHeight**, **TabFixedWidth**, and **MultiRow** properties.

You could set the tab control's  **Style** property to None if you wanted complete control over when a user could move between tabs. In prior versions of Microsoft Access, wizard dialogs were created by using multiple-page forms. You can now use a tab control create your own wizard with each page of the wizard contained on a separate page of a tab control with its **Style** property set to None.


## Example

The following example causes tabs to appear as buttons on the tab control named "TabCtl1" on the "Mailing List" form.


```vb
Forms("Mailing List").Controls("TabCtl1").Style = 1
```


## See also


#### Concepts


[TabControl Object](tabcontrol-object-access.md)

