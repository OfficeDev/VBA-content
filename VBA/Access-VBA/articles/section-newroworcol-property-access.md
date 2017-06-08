---
title: Section.NewRowOrCol Property (Access)
keywords: vbaac10.chm12193,vbaac10.chm4436
f1_keywords:
- vbaac10.chm12193,vbaac10.chm4436
ms.prod: access
api_name:
- Access.Section.NewRowOrCol
ms.assetid: c14c669a-9aff-6fc1-9bcc-b7bf011abed1
ms.date: 06/08/2017
---


# Section.NewRowOrCol Property (Access)

You can use the  **NewRowOrCol** property to specify whether a section and its associated data is printed in a new row or column within a multiple-column report or multiple-column form. Read/write **Byte**.


## Syntax

 _expression_. **NewRowOrCol**

 _expression_ A variable that represents a **Section** object.


## Remarks

The  **NewRowOrCol** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|None|0|(Default) The row or column breaks are determined by the settings in the  **Page Setup** dialog box (available by clicking **Page Setup** on the **File** menu) and the available space on the page.|
|Before Section|1|Microsoft Access starts printing the current section (the section for which you're setting the property, such as a group header section) in a new row or column. It then prints the next section, such as a detail section, in that same row or column.|
|After Section|2|Microsoft Access starts printing the current section, such as a group header section, in the current row or column. It starts printing the next section, such as a detail section, in the next row or column.|
|Before &; After|3|Microsoft Access starts printing the current section in a new row or column. It starts printing the following section in the next row or column.|
You can set the  **NewRowOrCol** property only in form Design view or report Design view.

The following items are some sample settings for a group header section in a multiple-column report. Make sure  **Down, Then Across** is selected under **Column Layout** on the **Columns** tab of the **Page Setup** dialog box.



|**Sample setting**|**Result**|
|:-----|:-----|
|Before Section|The group header is printed at the top of a new column.|
|After Section|The detail section is printed at the top of a new column.|
|Before &; After|The group header is printed in a column by itself, and the detail section is printed at the top of a new column.|
Sections in a form or report are normally printed vertically down a page. The default  **Column Layout** option is **Across, then Down**. You can print the sections in multiple columns across a page by clicking **Down, then Across** under **Column Layout** on the **Columns** tab of the **Page Setup** dialog box.

If you set the  **NewRowOrCol** property to Before Section, the vertical or horizontal orientation of the page affects how the section appears when printed. If you click **Across, then Down** under **Column Layout** on the **Columns** tab of the **Page Setup** dialog box, Microsoft Access starts printing the section at the beginning of a new row; if you click **Down, then Across**, Microsoft Access starts printing the section at the beginning of a new column.

The following example returns the  **NewRowOrCol** property setting and assigns it to the `intGetVal` variable.




```vb
Dim intGetVal As Integer 
intGetVal = Me.Section(1).NewRowOrCol
```

The next example presents two layouts for a report that divides data into four groups (Head1 to Head4). Each group includes three to six records, and each record has field a and field b. The layouts differ only in their settings under  **Column Layout** on the **Columns** tab of the **Page Setup** dialog box and the values of their **NewRowOrCol** properties. Note that the **Width** box under **Column Size** on the **Columns** tab must be set to the actual width of the field. Also, the Before Section setting of the **NewRowOrCol** property requires a page header section greater than zero for the **Down, then Across** option to function correctly.


|||
|:-----|:-----|
|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><b>Column Layout</b>  ? <b>Across, then Down</b><br /><br /> 
Head1 
<br />1a 1b     2a 2b     3a 3b     4a 4b<br /> 
5a 5b<br /> 
Head2<br /> 
1a 1b     2a 2b     3a 3b     4a 4b<br /> 
Head3<br /> 
1a 1b     2a 2b     3a 3b<br /> 
Head4<br /> 
1a 1b     2a 2b     3a 3b     4a 4b<br /> 
5a 5b     6a 6b<br /></p></li><li><p><b>Grid Settings</b>  ? <b>Number of Columns</b>  set to 4<br /><br /></p></li><li><p><b>NewRowOrCol</b>  property setting for group header section ? Before &; After</p></li></ul>|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><b>Column Layout</b>  ? <b>Down, then Across</b><br /><br />Head1     Head2     Head3     Head4<br /> 
1a 1b      1a 1b       1a 1b      1a 1b<br /> 
2a 2b      2a 2b       2a 2b      2a 2b<br /> 
3a 3b      3a 3b       3a 3b      3a 3b<br /> 
4a 4b      4a 4b       4a 4b<br /> 
5a 5b      5a 5b<br /> 
6a 6b<br /></p></li><li><p><b>Grid Settings</b>  ? <b>Number of Columns</b>  set to 4<br /><br /></p></li><li><p><b>NewRowOrCol</b>  property setting for group header section ? Before Section</p></li></ul>|

## See also


#### Concepts


[Section Object](section-object-access.md)

