---
title: Section.ForceNewPage Property (Access)
keywords: vbaac10.chm12192
f1_keywords:
- vbaac10.chm12192
ms.prod: access
api_name:
- Access.Section.ForceNewPage
ms.assetid: c523159f-f1f4-22b0-1aa3-05b7b213229a
ms.date: 06/08/2017
---


# Section.ForceNewPage Property (Access)

You can use the  **ForceNewPage** property to specify whether form sections detail, footer) or report sections (header, detail, footer) print on a separate page, rather than on the current page. Read/write **Byte**.


## Syntax

 _expression_. **ForceNewPage**

 _expression_ A variable that represents a **Section** object.


## Remarks

For example, you may have designed the last page of a report as an order form. If the report footer's  **ForceNewPage** property is set to Before Section, the order form is always printed on a new page.


 **Note**  The  **ForceNewPage** property does not apply to page headers or page footers.

The  **ForceNewPage** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|None|0|(Default) The current section (the section for which you're setting the property) is printed on the current page.|
|Before Section|1|The current section is printed at the top of a new page.|
|After Section|2|The section immediately following the current section is printed at the top of a new page.|
|Before &; After|3|The current section is printed at the top of a new page, and the next section is printed at the top of a new page.|
Here are some examples of the  **ForceNewPage** property setting.



|**Section**|**Sample setting**|**Description**|
|:-----|:-----|:-----|
|A group header displaying the year|Before Section|The group header is printed at the top of the page, followed by the detail section, group footer, and page footer.|
|A report detail section|After Section|The group footer is printed at the top of a new page.|
|A report header containing the report title and company logo.|After Section|The report title and logo are printed on a separate page at the beginning of the report.|

## Example

The following example returns the  **ForceNewPage** property setting for the detail section of the Sales By Date report and assigns it to the `intGetVal` variable.


```vb
Dim intGetVal As Integer 
intGetVal = Reports![Sales By Year].Section(acDetail).ForceNewPage
```


## See also


#### Concepts


[Section Object](section-object-access.md)

