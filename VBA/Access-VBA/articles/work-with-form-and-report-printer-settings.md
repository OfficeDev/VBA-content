---
title: Work with Form and Report Printer Settings
ms.prod: access
ms.assetid: 14a8aa00-9ad8-60f7-e103-791ab08c0e9e
ms.date: 06/08/2017
---


# Work with Form and Report Printer Settings

The  **[Printer](printer-object-access.md)** objects associated with **[Form](form-object-access.md)** and **[Report](report-object-access.md)** objects support the same properties and programming techniques as the **[Application](application-object-access.md)** object's **Printer** object. Use the **Printer** object of a **Form** or **Report** object when you want to set or retrieve printer settings for a specific form or report. You can change form and report printer settings temporarily, or you can save those settings with the form or report.


## Saving Printer Settings with a Form or Report

Whether a form or report uses the settings of the default application printer (the settings managed with the  **Application** object's **Printer** object) is determined by whether the form or report has previously saved printer settings. Printer settings for a form or report can be saved two ways:


- A user can save printer settings by opening the form or report in any view, and using the  **Print** or **Page Setup** dialog boxes to change the settings for the form or report.
    
- You can make changes to the  **Printer** object of a form or report in code, and those changes will be saved with the form or report if you use the **[Save](docmd-save-method-access.md)** method before closing the form or report, or specify **acSaveYes** for the _Save_ argument when using the **[Close](docmd-close-method-access.md)** method to close the form or report.
    

 **Note**  When printer settings are saved with a form or report, Access creates a new data structure for the form or report to contain the saved settings. Initially, this new data structure contains a copy of all of the settings of the default printer. Any settings the user or your code overrides are saved with the data structure. Access does not maintain any sort of inheritance between settings of the default printer and the settings saved with a form or report. If you change settings of the default printer after saving settings for a form or report, the settings that were originally saved will remain in effect.


## Determining Whether a Form or Report Has Saved Printer Settings

To determine whether a form or report has saved printer settings, you can read the  **UseDefaultPrinter** property of a **Form** or **Report** object using the following syntax:


```
expression .UseDefaultPrinter 

```

Where  _expression_ is any expression that returns a **Form** or **Report** object. The **UseDefaultPrinter** property is read/write in Design view and read-only in all other views.


## Clearing Saved Printer Settings

You can also use the  **UseDefaultPrinter** property like a method to clear saved settings from a form or report by setting its value to **True**. This is equivalent to opening the **Page Setup** dialog box for the form or report and selecting **Default Printer** on the **Page** tab.

You can set the  **UseDefaultPrinter** property only when a form or report is open in Design view. The following code fragment opens each of the reports in the current project and clears any report that has saved settings.




```vb
For Each obj In CurrentProject.AllReports 
    DoCmd.OpenReport ReportName:=obj.Name, View:=acViewDesign 
    If Not Reports(obj.Name).UseDefaultPrinter Then 
        Reports(obj.Name).UseDefaultPrinter = True 
        DoCmd.Save ObjectType:=acReport, ObjectName:=obj.Name 
    End If 
    DoCmd.Close 
Next obj 

```


## Preserving Form and Report Printer Settings

When you programmatically change printer property settings for forms or reports while the object is in any view other than Design view, those changes are automatically saved if the user interactively closes the form or report. The following procedure demonstrates how to save and restore a report's printer settings.


```vb
Sub RestoreReportPrinter() 
    Dim rpt As Report 
    Dim prtOld As Printer 
    Dim prtNew As Printer 
 
    ' Open the Invoice report in Print Preview. 
    DoCmd.OpenReport ReportName:="Invoice", View:=acViewPreview 
 
    ' Initialize rpt variable. 
    Set rpt = Reports!Invoice 
 
    ' Save the report's current printer settings 
    ' in the prtOld variable. 
    Set prtOld = rpt.Printer 
 
    ' Load the report's current printer settings 
    ' into the prtNew variable. 
    Set prtNew = rpt.Printer 
 
    ' Change the report's Orientation property. 
    prtNew.Orientation = acPRORLandscape 
 
    ' Change other Printer properties, and then print 
    ' or perform other operations here. 
 
    ' If you comment out the following line of code, 
    ' and a user interactively closes the report preview 
    ' any changes made to properties of the report's Printer 
    ' object are saved when the report is closed.  
    Set rpt.Printer = prtOld 
 
    ' Close report without saving. 
    DoCmd.Close ObjectType:=acReport, ObjectName:="Invoice", Save:=acSaveNo 
 
End Sub
```


