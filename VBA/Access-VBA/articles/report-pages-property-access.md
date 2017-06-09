---
title: Report.Pages Property (Access)
keywords: vbaac10.chm13722
f1_keywords:
- vbaac10.chm13722
ms.prod: access
api_name:
- Access.Report.Pages
ms.assetid: b97a6878-0a2c-3834-8f3d-6f4460dab3bd
ms.date: 06/08/2017
---


# Report.Pages Property (Access)

You can use the  **Pages** property to return information needed to print page numbers in a report. Read/write **Integer**.


## Syntax

 _expression_. **Pages**

 _expression_ A variable that represents a **Report** object.


## Remarks

This property is only available in Print Preview or when printing.

To refer to the  **Pages** property in a macro or Visual Basic, the form or report must include a text box whose **ControlSource** property is set to an expression that uses **Pages**. For example, you can use the following expressions as the **ControlSource** property setting for a text box in a page footer.



|**This expression**|**Prints**|
|:-----|:-----|
|=Page|A page number (for example, 1, 2, 3) in the page footer.|
|="Page " &; Page &; " of " &; Pages|"Page  _n_ of _nn_ " (for example, Page 1 of 5, Page 2 of 5) in the page footer.|
|=Pages|The total number pages in the report (for example, 5).|

## Example

The following example displays a message that tells how many pages the report contains. For this example to work, the report must include a text box for which the  **ControlSource** property is set to the expression `=Pages`. To test this example, paste the following code into the Page Event for the Alphabetical List of Products form.


```vb
Dim intTotalPages As Integer 
Dim strMsg As String 
 
intTotalPages = Me.Pages 
strMsg = "This report contains " &; intTotalPages &; " pages." 
MsgBox strMsg
```


## See also


#### Concepts


[Report Object](report-object-access.md)

