---
title: Label.HyperlinkAddress Property (Access)
keywords: vbaac10.chm10192
f1_keywords:
- vbaac10.chm10192
ms.prod: access
api_name:
- Access.Label.HyperlinkAddress
ms.assetid: ed50cbbe-f0bb-d096-2c50-920ad2f48eb9
ms.date: 06/08/2017
---


# Label.HyperlinkAddress Property (Access)

You can use the  **HyperlinkAddress** property to specify or determine the path to an object, document, Web page or other destination for a hyperlink associated with a the specified control. Read/write **String**.


## Syntax

 _expression_. **HyperlinkAddress**

 _expression_ A variable that represents a **Label** object.


## Remarks

The  **HyperlinkAddress** property is a string expression representing the path to a file (UNC path) or Web page (URL).


 **Note**  When you set the  **HyperlinkAddress** property, you automatically specify the **Address** property for the **Hyperlink** object associated with the control.


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


[Label Object](label-object-access.md)

