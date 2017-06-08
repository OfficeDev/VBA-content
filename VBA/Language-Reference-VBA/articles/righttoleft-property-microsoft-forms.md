---
title: RightToLeft Property (Microsoft Forms)
keywords: fm20.chm5282668
f1_keywords:
- fm20.chm5282668
ms.prod: office
ms.assetid: 2bd069aa-dd3a-c764-5b6c-6d49d381bd5c
ms.date: 06/08/2017
---


# RightToLeft Property (Microsoft Forms)



Specifies whether a given form supports bidirectional characteristics.
 **Syntax**
 _expression_**.RightToLeft** [= _value_ ]
 **Remarks**
 **RightToLeft** is a new property of the Microsoft Forms 2.0 form that the specified control is placed on; it's not a property of the control. Note, however, that **RightToLeft** affects both the form and any controls that are placed on it. The following table describes the two possible settings for this property.


|**Setting**|**Value**|
|:-----|:-----|
|False|0|
|True|1|

Microsoft Forms 2.0 controls that have the ability to exhibit bidirectional characteristics do so when the  **RightToLeft** property of the form is set to **True**. When this property is set to **False**, forms and controls do not exhibit bidirectional characteristics. Bidirectional features of the **RightToLeft** property are listed in the following table.


|**Microsoft Forms 2.0 components**|**RightToLeft = True behavior**|
|:-----|:-----|
|All forms and controls|The vertical scroll bar (if present) is on the left side of the form or control. The initial thumb location of the horizontal scroll bar is the rightmost position, and the rightmost portion of information is displayed.|
|Form|The title bar caption is right-aligned.|
|Frame control|The caption is placed in the upper-right corner of the frame.|
|TabStrip control and MultiPage control|Tabs are displayed in order from right to left for specific  **TabOrientation** settings (that is, the first tab begins at the right boundary of the control). For **TabOrientation** = **Top**, tabs are displayed from right to left along the top edge of the control. For **TabOrientation** = **Bottom**, tabs are displayed from right to left along the bottom edge of the control.<table><tr><th>**Note**</th></tr><tr><td>If the total width of all tabs exceeds the width of the control and  **MultiRow** = **False**, tabs are displayed from right to left as if **MultiRow** = **True**.</td></tr></table>|

|**Important**|
|:-----|  
|If you add Microsoft Forms 2.0 controls to a container surface that's not a Microsoft Forms 2.0 form (for example, if you add a  **TabStrip** control directly to an Arabic Word document), the control will behave as if the container has a **RightToLeft** property that's set to **False**.|

|**Note**|
|:-----|  
|The  **VerticalScrollBarSide** property described in Help has been replaced by the **RightToLeft** property described in this topic. Keep in mind that although you can use references to **VerticalScrollBarSide** in Visual Basic for Applications statements, the property is inoperative.|



