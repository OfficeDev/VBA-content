---
title: Application.CreateReport Method (Access)
keywords: vbaac10.chm12517
f1_keywords:
- vbaac10.chm12517
ms.prod: access
api_name:
- Access.Application.CreateReport
ms.assetid: 4b086f8c-8017-0b5f-72a7-7c180c32f52d
ms.date: 06/08/2017
---


# Application.CreateReport Method (Access)

The  **CreateReport** method creates a report and returns a **[Report](report-object-access.md)** object. For example, suppose you are building a custom wizard to create a sales report. You can use the **CreateReport** method in your wizard to create a new report based on a specified report template.


## Syntax

 _expression_. **CreateReport**( ** _Database_**, ** _ReportTemplate_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Database_|Optional|**Variant**|The name of the database that contains the report template you want to use to create a report. If you want the current database, omit this argument. If you want to use an open library database, specify the library database with this argument.|
| _ReportTemplate_|Optional|**Variant**| The name of the report you want to use as a template to create a new report.|

### Return Value

Report


## Remarks

You can use the  **CreateReport** method when designing a wizard that creates a new report.

The  **CreateReport** method open a new, minimized report in report Design view.

If the name you use for the  _reporttemplate_ argument isn't valid, Visual Basic uses the report template specified by the **Report Template** setting on the **Forms/Reports** tab of the **Options** dialog box.


## Example

The following example creates a report in the current database by using the template specified by the  **Report Template** setting on the **Forms/Reports** tab of the **Options** dialog box.


```vb
Sub NormalReport() 
 Dim rpt As Report 
 
 Set rpt = CreateReport ' Create minimized report. 
 DoCmd.Restore ' Restore report. 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

