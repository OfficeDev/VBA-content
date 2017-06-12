---
title: Hyperlink.EmailSubject Property (Access)
keywords: vbaac10.chm10118
f1_keywords:
- vbaac10.chm10118
ms.prod: access
api_name:
- Access.Hyperlink.EmailSubject
ms.assetid: e2854e40-d16c-f854-3543-80fc14c8f728
ms.date: 06/08/2017
---


# Hyperlink.EmailSubject Property (Access)

You can use the  **EmailSubject** property to specify or determine return the email subject line of a hyperlink to an object, document, Web page or other destination for a command button, image control, or label control. Read/write **String**.


## Syntax

 _expression_. **EmailSubject**

 _expression_ A variable that represents a **Hyperlink** object.


## Remarks

When you move the cursor over a command button, image control, or label control whose  **HyperlinkAddress** property is set, the cursor changes to an upward-pointing hand. Clicking the control displays the object or Web page specified by the link.

To open objects in the current database, leave the  **HyperlinkAddress** property blank and specify the object type and object name you want to open in the **HyperlinkSubAddress** property by using the syntax " _objecttype objectname_". If you want to open an object contained in another Microsoft Access database, enter the database path and file name in the  **HyperlinkAddress** property and specify the database object to open by using the **HyperlinkSubAddress** property.

The  **HyperlinkAddress** property can contain an absolute or a relative path to a target document. An absolute path is a fully qualified URL or UNC path to a document. A relative path is a path related to the base path specified in the **Hyperlink Base** setting in the _DatabaseName_ **Properties** dialog box (available by clicking **Database Properties** on the **File** menu) or to the current database path. If Microsoft Access can't resolve the **HyperlinkAddress** property setting to a valid URL or UNC path, it will assume you've specified a path relative to the base path contained in the **Hyperlink Base** setting or the current database path.


 **Note**  When you follow a hyperlink to another Microsoft Access database object, the database Startup properties are applied. For example, if the destination database has a Display form set, that form is displayed when the database opens.


 **Note**  When you create a hyperlink by using the  **Insert Hyperlink** dialog box, Microsoft Access automatically sets the **EmailSubject** property to the location specified in the **Subject** box of the **E-Mail Address** tab.


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-access.md)

