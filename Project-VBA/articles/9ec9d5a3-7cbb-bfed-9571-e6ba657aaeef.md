
# Application.GanttBarFormatEx Method (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Changes the formatting of Gantt bars from their default styles, where colors can be hexadecimal RGB values.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **GanttBarFormatEx**( **_TaskID_**,  **_GanttStyle_**,  **_StartShape_**,  **_StartType_**,  **_StartColor_**,  **_MiddleShape_**,  **_MiddlePattern_**,  **_MiddleColor_**,  **_EndShape_**,  **_EndType_**,  **_EndColor_**,  **_LeftText_**,  **_RightText_**,  **_TopText_**,  **_BottomText_**,  **_InsideText_**,  **_Reset_**,  **_ProjectName_**)

 _expression_An expression that returns an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|TaskID|Optional| **Long**|The identification number of the task to be changed on the Gantt chart. The default is to change the Gantt bars of the selected tasks.|
|GanttStyle|Optional| **Integer**|The style applied to the Gantt bar to be formatted. The value for GanttStyle is based on the position of the bar style in the  **Bar Styles** dialog box. For example, the value 3 returns the third bar style in the list.|
|StartShape|Optional| **Integer**|The start shape of the Gantt bar. Can be one of the  ** [PjBarEndShape](0598711b-46ad-1940-103b-12345f32efd8.md)** constants.|
|StartType|Optional| **Integer**|The start type of the Gantt bar. Can be one of the  ** [PjBarType](abc6a0b2-90bd-48d4-283a-a53618856692.md)** constants.|
|StartColor|Optional| **Long**|The color of the start shape of the Gantt bar. Can be a hexadecimal RGB value, where red is the last byte. For example, &amp;H00FFFF is yellow. |
|MiddleShape|Optional| **Integer**|The middle shape of the Gantt bar. Can be one of the  ** [PjBarShape](057356dc-9cab-fbdc-563e-f81cc54a2c33.md)** constants.|
|MiddlePattern|Optional| **Integer**|The middle pattern of the Gantt bar. Can be one of the  ** [PjFillPattern](4f6af32c-5efd-42b6-4017-20a1497c1b6d.md)** constants.|
|MiddleColor|Optional| **Long**|The color of the middle section Gantt bar. Can be a hexadecimal RGB value, where red is the last byte. For example, &amp;HFF00FF is purple. |
|EndShape|Optional| **Integer**|The end shape of the Gantt bar. Can be one of the  **PjBarEndShape** constants.|
|EndType|Optional| **Integer**|The end type of the Gantt bar. Can be one of the following  **PjBarType** constants: **pjDashed**,  **pjFramed**, or  **pjSolid**.|
|EndColor|Optional| **Long**|The color of the end shape of the Gantt bar. Can be a hexadecimal RGB value, where red is the last byte. For example, &amp;HFFFF00 is blue-green. |
|LeftText|Optional| **String**|The task field to display to the left of the Gantt bar.|
|RightText|Optional| **String**|The task field to display to the right of the Gantt bar.|
|TopText|Optional| **String**|The task field to display above the Gantt bar.|
|BottomText|Optional| **String**|The task field to display below the Gantt bar.|
|InsideText|Optional| **String**|The task field to display inside the Gantt bar.|
|Reset|Optional| **Boolean**| **True** if the bar formatting is reset to the default formatting of the style in the **Bar Styles** dialog box; otherwise, **False**.|
|ProjectName|Optional| **String**|The name of the project containing  **TaskID** if consolidation is involved. The default value is the name of the active project.|

### Return Value

 **Boolean**


## Remarks
<a name="sectionSection1"> </a>

Using the  **GanttBarFormatEx** method without specifying any arguments displays the **Format Bar** dialog box.

 To define the default styles where colors can be hexadecimal RGB values, use the ** [GanttBarEditEx](b574b975-a869-31ba-e525-df8775330b0a.md)** method.


## Example
<a name="sectionSection2"> </a>

The following example displays a medium red diamond shape for the start of the task with the Task ID of 3.


```
Sub GanttBar_Format() 
 
    'Activate Gantt Chart view 
    ViewApply Name:="&amp;Gantt Chart" 
    GanttBarFormatEx TaskID:=3, StartShape:=pjDiamond, StartType:=pjSolid, StartColor:=&amp;H8888FF
End Sub
```

