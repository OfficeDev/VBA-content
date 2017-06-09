---
title: Overview of the IRibbonUI Object
ms.prod: outlook
ms.assetid: ef273431-550f-4ff6-b964-79d05b09bea5
ms.date: 06/08/2017
---


# Overview of the IRibbonUI Object

An add-in can use the  [IRibbonUI](http://msdn.microsoft.com/library/d323aa21-de74-e821-c914-db71ef3b9c5e%28Office.15%29.aspx) object to invalidate controls or menu items, and to update their content in the corresponding Microsoft Outlook user interface. The add-in specifies callback methods in the XML that [IRibbonExtensibility.GetCustomUI](http://msdn.microsoft.com/library/a0106415-999e-94da-379c-70fb7aa6119f%28Office.15%29.aspx) returns. These callback methods handle events for custom controls or custom menu items. 

When Outlook calls one of these methods, it passes an **IRibbonUI** object as a parameter to the callback method. The **IRibbonUI** object is scoped so that the add-in can only invalidate its own controls or menu items that use the object. The add-in cannot invalidate the controls or menu items that another add-in created.

 **IRibbonUI** exposes the following methods to customize the user interface in Outlook:


| **Method**| **Action**| **Description**|
|:-----|:-----|:-----|
| **[Invalidate()](http://msdn.microsoft.com/library/068cd459-76c2-b1d3-ed7d-50fa88c4db73%28Office.15%29.aspx)**|Callback|Marks all of the custom controls or menu items in your add-in for update.|
| **[InvalidateControl(string controlID)](http://msdn.microsoft.com/library/33af7933-66f7-51e9-895e-07a6222973d2%28Office.15%29.aspx)**|Callback|Marks a specific control or menu item that is defined by a  _controlID_ in your add-in for update.|
| **[ActivateTab](http://msdn.microsoft.com/library/32f5205c-6ab1-e3a6-6bae-5f36706c4d0d%28Office.15%29.aspx)**|Callback|Activates the specified custom tab on the Microsoft Office Fluent ribbon.|
| **[ActivateTabQ](http://msdn.microsoft.com/library/bf664b52-2660-2ce7-a01b-83b459f66e09%28Office.15%29.aspx)**|Callback|Activates the specified custom tab on the ribbon by using the fully qualified name of the tab.|
To minimize the impact on performance, use the  **InvalidateControl** method instead of the **Invalidate** method unless you actually need to invalidate all the custom controls or menu items that your add-in defines. Calling **Invalidate** invalidates all controls and menu items that your add-in defines, and callbacks occur on open explorers, inspectors, and menus.

## See also


#### Concepts


 [Implementing the IRibbonExtensibility Interface](implementing-the-iribbonextensibility-interface.md)<br>
 [Office Fluent User Interface Extensibility for Outlook](office-fluent-user-interface-extensibility-for-outlook.md)<br>

