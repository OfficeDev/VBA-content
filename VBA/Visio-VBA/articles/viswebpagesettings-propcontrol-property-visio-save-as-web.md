---
title: VisWebPageSettings.PropControl Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.PropControl
ms.assetid: 615e5038-d84d-9527-6987-95f289da77d9
ms.date: 06/08/2017
---


# VisWebPageSettings.PropControl Property (Visio Save As Web)

Determines whether a shape's custom properties (shape data) are displayed in the Web page. Read/write.


## Syntax

 _expression_. **PropControl**

 _expression_An expression that returns a  ** [VisWebPageSettings](http://msdn.microsoft.com/library/14280ea7-e8b1-d4b2-941b-121f2c17f787%28Office.15%29.aspx)** object.


### Return Value

 **Long**


## Remarks

 **PropControl** returns non-zero ( **True**) if custom properties are set to be displayed in the Web page; otherwise, it returns zero ( **False**). The default is  **True**.

Set  **PropControl** to a non-zero value ( **True**) to display custom properties in the Web page; otherwise, set it to zero ( **False**). 

If you choose to display custom properties, a  **Custom Properties** control appears in the left frame in the browser window, displaying custom properties (shape data) that are associated with a shape when you press CTRL and click the shape.

If a shape is part of a group, and both the group and its subshapes have custom properties, the custom properties are displayed in the browser according to the behavior defined in the  **Selection** list box on the **Behavior** dialog box (with the group shape selected, click **Behavior** on the **Format** menu).

The selected behavior determines the display as follows: 


- With  **Group only** or **Group first**, Save as Web Page displays the group's custom properties.
    
- With  **Members first**, Save as Web Page displays the subshape's custom properties when the mouse pointer moves over a subshape that has custom properties and group custom properties for those subshapes that do not have custom properties.
    


This behavior can also be set in the SelectMode cell in the Group Properties section of the group shape in the Visio ShapeSheet Spreadsheet.

The value of the  **PropControl** property corresponds to the setting of the **Details** check box in the **Publishing options** list on the **General** tab of the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, and then click  **Publish**).


