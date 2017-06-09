---
title: Application.CalendarBarStylesEditEx Method (Project)
keywords: vbapj.chm2146
f1_keywords:
- vbapj.chm2146
ms.prod: project-server
api_name:
- Project.Application.CalendarBarStylesEditEx
ms.assetid: 3b7cb188-fff6-b9c1-a673-34774791c043
ms.date: 06/08/2017
---


# Application.CalendarBarStylesEditEx Method (Project)

Changes the style of the specified type of bar in the Calendar view, where  _Color_ can be a hexadecimal value.


## Syntax

 _expression_. **CalendarBarStylesEditEx**( ** _Item_**, ** _Bar_**, ** _Pattern_**, ** _Color_**, ** _Align_**, ** _Wrap_**, ** _Shadow_**, ** _Field1_**, ** _Field2_**, ** _Field3_**, ** _Field4_**, ** _Field5_**, ** _SplitPattern_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required|**Long**|The type of calendar bar style to edit. Can be one of the following  **[PjBarItem](pjbaritem-enumeration-project.md)** constants: **pjBarNonCritical**, **pjBarCritical**, **pjBarSummary**, **pjBarMilestone**, **pjBarMarked**, **pjBarHighlighted**, **pjBarProjectSummary**, or **pjBarExternalTask**.|
| _Bar_|Optional|**Long**|The bar type. Can be one of the following  **[PjCalendarBarType](pjcalendarbartype-enumeration-project.md)** constants: **pjNormalBar**, **pjLineBar**, or **pjNoBar**.|
| _Pattern_|Optional|**Long**|The bar pattern. Can be one of the  **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|
| _Color_|Optional|**Long**|The bar color. Can be a hexadecimal value for the RGB color, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow.|
| _Align_|Optional|**Long**|The justification of text in the bar. Can be one of the following  **[PjAlignment](pjalignment-enumeration-project.md)** constants: **pjLeft**, **pjCenter**, or **pjRight**.|
| _Wrap_|Optional|**Boolean**|**True** if Project wraps text in the bar; otherwise, **False**.|
| _Shadow_|Optional|**Boolean**|**True** if the bar has a shadow; otherwise, **False**.|
| _Field1_|Optional|**String**|The first field to display in the bar.|
| _Field2_|Optional|**String**|The second field to display in the bar.|
| _Field3_|Optional|**String**|The third field to display in the bar.|
| _Field4_|Optional|**String**|The fourth field to display in the bar.|
| _Field5_|Optional|**String**|The fifth field to display in the bar.|
| _SplitPattern_|Optional|**Long**|The line pattern used to display split tasks. Can be one of the following  **[PjLineType](pjlinetype-enumeration-project.md)** constants: **pjNoLines**, **pjDash**, **pjCloseDot**, **pjContinuous**, or **pjDot**.|

### Return Value

 **Boolean**


## Remarks

Specifying a value for any of  _Field1_ through _Field5_ requires that all preceding _Field_ arguments also be specified. For example, specifying _Field3_ also requires _Field1_ and _Field2_ to be specified.


 **Note**  The  _Field1_ to _Field5_ parameters cannot use the **PjFields** constants. To see the field names that you can add to calendar bars, open the Calendar view, click the **Format** tab in the **Calendar Tools** group. Click **Bar Styles** on the Ribbon, and then click the ** Field(s)** drop-down list.


## Example

The following example sets critical tasks as normal bars, the color to purple with diagonal stripes, and the fields to include the task name and assigned resource names. The example also sets summary tasks as line bars and the color to green.


```vb
Sub CalendarBar_StyleEdit()
    'Activate Caldender view 
    ViewApply Name:="Calendar" 
 
    CalendarBarStylesEditEx Item:=pjBarCritical, Bar:=PjCalendarBarType.pjNormalBar, _
          Color:=&;HAA00AA, Pattern:=PjFillPattern.pjDiagonalRightPattern, _
          Field1:="Name", Field2:="Resource Names" 
    CalendarBarStylesEditEx Item:=pjBarSummary, Bar:=PjCalendarBarType.pjLineBar, _
    Color:=&;H1AA01 
End Sub
```


 **Note**  If you use any of the  **PjColor** enumeration constants for the _Color_ parameter, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the **CalendarBarStylesEditEx** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the **[CalendarBarStylesEdit](application-calendarbarstylesedit-method-project.md)** method.


