---
title: OLE Programmatic Identifiers (Office)
keywords: vbaof11.chm5221270
f1_keywords:
- vbaof11.chm5221270
ms.prod: office
ms.assetid: e27f70fd-9e04-a8d0-d4e8-d57076ecf9b3
ms.date: 06/08/2017
---


# OLE Programmatic Identifiers (Office)

You can use an OLE programmatic identifier (sometimes called a ProgID) to create an Automation object. The following tables list OLE programmatic identifiers for ActiveX controls and the Office applications.

[ActiveX Controls](#activexcontrols)

[Microsoft Access](#access)

[Microsoft Excel](#excel)
[Microsoft Graph](#graph)
[Microsoft Outlook](#outlook)
[Microsoft PowerPoint](#powerpoint)
[Microsoft Word](#word)

## ActiveX Controls
<a name="activexcontrols"> </a>

To create the ActiveX controls that are listed in the following table, use the corresponding OLE programmatic identifier.



|**To create this control**|**Use this identifier**|
|:-----|:-----|
|CheckBox|Forms.CheckBox.1|
|ComboBox|Forms.ComboBox.1|
|CommandButton|Forms.CommandButton.1|
|Frame|Forms.Frame.1|
|Image|Forms.Image.1|
|Label|Forms.Label.1|
|ListBox|Forms.ListBox.1|
|MultiPage|Forms.MultiPage.1|
|OptionButton|Forms.OptionButton.1|
|ScrollBar|Forms.ScrollBar.1|
|SpinButton|Forms.SpinButton.1|
|TabStrip|Forms.TabStrip.1|
|TextBox|Forms.TextBox.1|
|ToggleButton|Forms.ToggleButton.1|

## Microsoft Access
<a name="access"> </a>

To create the Microsoft Access objects that are listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Access that is available on the computer where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|Application|Access.Application|
|CurrentData|Access.CodeData, Access.CurrentData|
|CurrentProject|Access.CodeProject, Access.CurrentProject|

## Microsoft Excel
<a name="excel"> </a>

To create the Microsoft Excel objects that are listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Excel that is available on the computer where the macro is running.



|**To create this object**|**Use this identifier**|**Comments**|
|:-----|:-----|:-----|
|Application|Excel.Application||
|Workbook|Excel.AddIn||
|Workbook|Excel.Chart|Returns a workbook that contains two worksheets; one for the chart and one for its data. The chart worksheet is the active worksheet.|
|Workbook|Excel.Sheet|Returns a workbook with one worksheet.|

## Microsoft Graph
<a name="graph"> </a>

To create the Microsoft Graph objects that are listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Graph that is available on the computer where the macro is running.



|**To create this object**|**Use this identifier**|
|:-----|:-----|
|Application|MSGraph.Application|
|Chart|MSGraph.Chart|

## Microsoft Outlook
<a name="outlook"> </a>

To create the Microsoft Outlook object that are listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Outlook that is available on the computer where the macro is running.



|**To create this object**|**Use this identifier**|
|:-----|:-----|
|Application|Outlook.Application|
To create the ActiveX controls that are specific to the Outlook forms listed in the following table, use the corresponding OLE programmatic identifier.



|**To create this Microsoft Office Outlook control**|**Use this identifier**|
|:-----|:-----|
|OlkBusinessCardControl|Outlook.OlkBusinessCardControl |
|OlkCategory|Outlook.OlkCategoryStrip|
|OlkCheckBox|Outlook.OlkCheckBox|
|OlkComboBox|Outlook.OlkComboBox|
|OlkCommandButton|Outlook.OlkCommandButton|
|OlkContactPhoto|Outlook.OlkContactPhoto|
|OlkDateControl|Outlook.OlkDateControl|
|OlkFrameHeader|Outlook.OlkFrameHeader|
|OlkInfoBar|Outlook.OlkInfoBar|
|OlkLabel|Outlook.OlkLabel|
|OlkListBox|Outlook.OlkListBox|
|OlkOptionButton|Outlook.OlkOptionButton|
|OlkPageControl|Outlook.OlkPageControl|
|OlkSenderPhoto|Outlook.OlkSenderPhoto|
|OlkTextBox|Outlook.OlkTextBox|
|OlkTimeControl|Outlook.OlkTimeControl|
|OlkTimeZoneControl|Outlook.OlkTimeZone|

## Microsoft PowerPoint
<a name="powerpoint"> </a>

To create the Microsoft PowerPoint object that are listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of PowerPoint that is available on the computer where the macro is running.



|**To create this object**|**Use this identifier**|
|:-----|:-----|
|Application|PowerPoint.Application|

## Microsoft Word
<a name="word"> </a>

To create the Microsoft Word objects that are listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Word that is available on the computer where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|Application|Word.Application|
|Document|Word.Document, Word.Template|
|Global|Word.Global|

