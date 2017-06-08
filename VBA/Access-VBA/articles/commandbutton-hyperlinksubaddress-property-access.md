---
title: CommandButton.HyperlinkSubAddress Property (Access)
keywords: vbaac10.chm10461
f1_keywords:
- vbaac10.chm10461
ms.prod: access
api_name:
- Access.CommandButton.HyperlinkSubAddress
ms.assetid: 1c8af1e0-f978-0eb2-c3b5-f5ea9ab84892
ms.date: 06/08/2017
---


# CommandButton.HyperlinkSubAddress Property (Access)

You can use the  **HyperlinkSubAddress** property to specify or determine a location within the target document specified by the **HyperlinkAddress** property. Read/write **String**.


## Syntax

 _expression_. **HyperlinkSubAddress**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **HyperlinkSubAddress** property can be an object within a Microsoft Access database, a bookmark within a Microsoft Word document, a named range within a Microsoft Excel spreadsheet, a slide within a Microsoft PowerPoint presentation, or a location within an HTML document.

The  **HyperlinkSubAddress** property is a string expression representing a named location within the target document specified by the **HyperlinkAddress** property.


 **Note**  When you create a hyperlink by using the  **Insert Hyperlink** dialog box, Microsoft Access automatically sets the **HyperlinkAddress** property and **HyperlinkSubAddress** to the location specified in the **Type the file or Web page name** box. The **HyperlinkSubAddress** property can also be set to the location specified in the **Select an object in this database** box.

If you copy a hyperlink from another application and paste it into a form or report, Microsoft Access creates a label control with its  **Caption** property, **HyperlinkAddress** property, and **HyperlinkSubAddress** property automatically set.

When you move the cursor over a command button, image control, or label control whose  **HyperlinkAddress** property is set, the cursor changes to an upward-pointing hand. Clicking the control displays the object or Web page specified by the link.

To open objects in the current database, leave the  **HyperlinkAddress** property blank and specify the object type and object name you want to open in the **HyperlinkSubAddress** property by using the syntax "objecttype objectname". If you want to open an object contained in another Microsoft Access database, enter the database path and file name in the **HyperlinkAddress** property and specify the database object to open by using the **HyperlinkSubAddress** property.

The  **HyperlinkAddress** property can contain an absolute or a relative path to a target document. An absolute path is a fully qualified URL or UNC path to a document. A relative path is a path related to the base path specified in the **Hyperlink Base** setting in the **Properties** dialog box (available by clicking **Database Properties** on the **File** menu) or to the current database path. If Microsoft Access can't resolve the **HyperlinkAddress** property setting to a valid URL or UNC path, it will assume you've specified a path relative to the base path contained in the **Hyperlink Base** setting or the current database path.


 **Note**  When you follow a hyperlink to another Microsoft Access database object, the database Startup properties are applied. For example, if the destination database has a Display form set, that form is displayed when the database opens.

The following table contains examples of  **HyperlinkAddress** and **HyperlinkSubAddress** property settings.



|**HyperlinkAddress**|**HyperlinkSubAddress**|**Description**|
|:-----|:-----|:-----|
|http://www.microsoft.com/||The Microsoft home page on the Web.|
|C:\Program Files\Microsoft Office\Office\Samples\Cajun.htm||The Cajun Delights page in the Access sample applications subdirectory.|
|C:\Program Files\Microsoft Office\Office\Samples\Cajun.htm|NewProducts|The "NewProducts" Name tag in the Cajun Delights page.|
|C:\Personal\MyResume.doc|References|The bookmark named "References" in the Microsoft Word document "MyResume.doc".|
|C:\Finance\First Quarter.xls|Sheet1!TotalSales|The range named "TotalSales" in the Microsoft Excel spreadsheet "First Quarter.xls".|
|C:\Presentation\NewPlans.ppt|10|The 10th slide in the Microsoft PowerPoint document "NewPlans.ppt".|

## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

