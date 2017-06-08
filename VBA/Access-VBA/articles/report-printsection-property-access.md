---
title: Report.PrintSection Property (Access)
keywords: vbaac10.chm13730
f1_keywords:
- vbaac10.chm13730
ms.prod: access
api_name:
- Access.Report.PrintSection
ms.assetid: 745f4624-557b-0a4c-d4f4-9f0ba4113a61
ms.date: 06/08/2017
---


# Report.PrintSection Property (Access)

The  **PrintSection** property specifies whether a section should be printed. Read/write **Boolean**.


## Syntax

 _expression_. **PrintSection**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **PrintSection** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|**True**|(Default) The section is printed.|
|**False**|The section isn't printed.|

 **Note**  To set this property, specify a macro or event procedure for a section's  **[OnFormat](section-onformat-property-access.md)** property.

Microsoft Access sets this property to  **True** before each section's **Format** event.


## Example

The following example does not print the section "PageHeaderSection" of the "Product Summary" report.


```vb
Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer) 
 
 Reports("Product Summary").PrintSection = False 
 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

