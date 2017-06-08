---
title: OLE Programmatic Identifiers (PowerPoint)
keywords: vbapp10.chm5193172
f1_keywords:
- vbapp10.chm5193172
ms.prod: powerpoint
ms.assetid: c0e766ee-09af-b20f-2eec-0c73ea1615a4
ms.date: 06/08/2017
---


# OLE Programmatic Identifiers (PowerPoint)

You can use an OLE programmatic identifier (sometimes called a ProgID) to create an Automation object. The following tables list OLE programmatic identifiers for ActiveX controls, Office applications, and Office Web Components.


## ActiveX Controls

To create the ActiveX controls listed in the following table, use the corresponding OLE programmatic identifier.



|**To create this control**|**Use this identifier**|
|:-----|:-----|
|**CheckBox**|Forms.CheckBox.1|
|**ComboBox**|Forms.ComboBox.1|
|**CommandButton**|Forms.CommandButton.1|
|**Frame**|Forms.Frame.1|
|**Image**|Forms.Image.1|
|**Label**|Forms.Label.1|
|**ListBox**|Forms.ListBox.1|
|**MultiPage**|Forms.MultiPage.1|
|**OptionButton**|Forms.OptionButton.1|
|**ScrollBar**|Forms.ScrollBar.1|
|**SpinButton**|Forms.SpinButton.1|
|**TabStrip**|Forms.TabStrip.1|
|**TextBox**|Forms.TextBox.1|
|**ToggleButton**|Forms.ToggleButton.1|

## Access

To create the Access objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Access available on the computer where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|**Application**|Access.Application|
|**CurrentData**|Access.CodeData, Access.CurrentData|
|**CurrentProject**|Access.CodeProject, Access.CurrentProject|
|**DefaultWebOptions**|Access.DefaultWebOptions|

## Excel

To create the Excel objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Excel available on the computer where the macro is running.



|**To create this object**|**Use one of these identifiers**|**Comments**|
|:-----|:-----|:-----|
|**Application**|Excel.Application||
|**Workbook**|Excel.AddIn||
|**Workbook**|Excel.Chart|Returns a workbook containing two worksheets; one for the chart and one for its data. The chart worksheet is the active worksheet.|
|**Workbook**|Excel.Sheet|Returns a workbook with one worksheet.|

## Microsoft Graph

To create the Microsoft Graph objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Graph available on the computer where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|**Application**|MSGraph.Application|
|**Chart**|MSGraph.Chart|

## Outlook

To create the Outlook object given in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Outlook available on the computer where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|**Application**|Outlook.Application|

## PowerPoint

To create the PowerPoint object given in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of PowerPoint available on the computer where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|**Application**|PowerPoint.Application|

## Word

To create the Word objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Word available on the computer where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
|**Application**|Word.Application|
|**Document**|Word.Document, Word.Template|
|**Global**|Word.Global|

