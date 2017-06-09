---
title: Application.OptionsEditEx Method (Project)
keywords: vbapj.chm2159
f1_keywords:
- vbapj.chm2159
ms.prod: project-server
api_name:
- Project.Application.OptionsEditEx
ms.assetid: d735d118-f004-ba67-7aa5-290ff256da10
ms.date: 06/08/2017
---


# Application.OptionsEditEx Method (Project)

Sets options for Project, where colors can be hexadecimal values, or opens the  **Project Options** dialog box.


## Syntax

 _expression_. **OptionsEditEx**( ** _MoveAfterReturn_**, ** _DragAndDrop_**, ** _UpdateLinks_**, ** _CopyResourceUsageHeader_**, ** _PhoneticInfo_**, ** _PhoneticType_**, ** _MinuteLabelDisplay_**, ** _HourLabelDisplay_**, ** _DayLabelDisplay_**, ** _WeekLabelDisplay_**, ** _YearLabelDisplay_**, ** _SpaceBeforeTimeLabel_**, ** _SetDefaults_**, ** _MonthLabelDisplay_**, ** _SetDefaultsTimeUnits_**, ** _HyperlinkColor_**, ** _FollowedHyperlinkColor_**, ** _UnderlineHyperlinks_**, ** _SetDefaultsHyperlink_**, ** _InCellEditing_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MoveAfterReturn_|Optional|**Boolean**|**True** if the next cell or field becomes active when ENTER is pressed. **False** if the current cell or field remains active. The **Move selection after enter** option is on the **Advanced** tab of the **Project Options** dialog box.|
| _DragAndDrop_|Optional|**Boolean**|**True** if cells may be copied or moved by dragging them; otherwise, **False**. The **Allow cell drag and drop** option is on the **Advanced** tab.|
| _UpdateLinks_|Optional|**Boolean**|**True** if Project asks to update automatic links when the relevant information changes; otherwise, **False**. The **Ask to update automatic links** option is on the **Advanced** tab.|
| _CopyResourceUsageHeader_|Optional||Due to changes in the Project object model, this argument is ignored. It is retained so that existing macros that use this argument do not cause errors.|
| _PhoneticInfo_|Optional|**Boolean**|**True** if phonetic information is automatically provided for resource names and custom fields; otherwise, **False**. The PhoneticInfo argument is ignored unless the Japanese version of Project is used.|
| _PhoneticType_|Optional|**Integer**|Specifies the type of characters used to display phonetic information. Can be one of the following  **[PjPhoneticType](pjphonetictype-enumeration-project.md)** constants: **pjKatakanaHalf**, **pjKatakana**, or **pjHiragana**. The PhoneticType argument is ignored unless the Japanese version of Project is used.|
| _MinuteLabelDisplay_|Optional|**Integer**|Specifies how the minute label displays. The minute label display corresponds to the  **Minutes** list, which is on the **Advanced** tab. For example, setting the MinuteLabelDisplay argument to 0 sets the **Minutes** list to the first value in the list ( **m**). Valid values are 0?2.|
| _HourLabelDisplay_|Optional|**Integer**|Specifies how the hour label displays. The hour label display corresponds to the  **Hours** list, which is found on the **Advanced** tab. For example, setting the HourLabelDisplay argument to 2 sets the **Hours** list to the the third value in the list ( **hour**). Valid values are 0?2.|
| _DayLabelDisplay_|Optional|**Integer**|Specifies how the day label displays. The day label display corresponds to the  **Days** list, which is found on the **Advanced** tab. For example, setting the DayLabelDisplay argument to 1 sets the **Days** list to the second value in the list ( **dy**). Valid values are 0?2.|
| _WeekLabelDisplay_|Optional|**Integer**|Specifies how the week label displays. The week label display corresponds to the  **Weeks** list, which is found on the **Advanced** tab. For example, setting the WeekLabelDisplay argument to 0 sets the **Weeks** list to the first value in the list ( **w**). Valid values are 0?2.|
| _YearLabelDisplay_|Optional|**Integer**| Specifies how the year label displays. The year label display corresponds to the **Years** list, which is found on the **Advanced** tab. For example, setting the YearLabelDisplay argument to 1 sets the **Years** list to the second value in the list ( **yr**). Valid values are 0?2.|
| _SpaceBeforeTimeLabel_|Optional|**Boolean**|**True** if a time value should be separated from the time label by a space; otherwise, **False**. Corresponds to the **Add space before label** option on the **Advanced** tab.|
| _SetDefaults_|Optional|**Boolean**|**True** if values of the arguments in the **OptionsEdit** method are set to the default values for new projects. The default value is **False**, which means that options are set only for the active project.|
| _MonthLabelDisplay_|Optional|**Integer**|Specifies how the month label displays. The month label display corresponds to the  **Months** list, which is found on the **Advanced** tab. For example, setting the **MonthLabelDisplay** property to 2 sets the **Months** list to the third value in the list ( **month**). Valid values are 0?2.|
| _SetDefaultsTimeUnits_|Optional|**Boolean**|**True** if the values of time units specified in the **Display options for this project** section ( **Advanced** tab) are used as the default values for new projects. The default value is **False**, which means that the display options of time units are set only for the active project.|
| _HyperlinkColor_|Optional|**Integer**|The color used to denote unfollowed hyperlinks. Can be a hexadecimal RBG value, where red is the last byte. For example, &;HFF0000 is blue.|
| _FollowedHyperlinkColor_|Optional|**Long**|The color used to denote followed hyperlinks. Can be a hexadecimal RBG value, where red is the last byte. For example, &;HFF00FF is purple.|
| _UnderlineHyperlinks_|Optional|**Boolean**|**True** if hyperlinks are underlined; otherwise, **False**.|
| _SetDefaultsHyperlink_|Optional|**Boolean**|**True** if the hyperlink options specified in the **Display options for this project**section ( **Advanced** tab) are used as the default values for new projects. The default value is **False**, which means that the hyperlink options are set only for the active project.|
| _InCellEditing_|Optional|**Boolean**|**True** if in-cell editing is enabled; otherwise, **False**. Corresponds to the **Edit directly in cell** option in the **Edit** section of the **Advanced** tab.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, its default value is specified by the current setting in the  **Project Options** dialog box.

Using the  **OptionsEditEx** method with no arguments displays the **Project Options** dialog box with the **General** tab selected.


