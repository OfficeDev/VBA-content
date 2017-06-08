---
title: Close a Report Automatically If It Does Not Contain Any Records
ms.prod: access
ms.assetid: 9b160bd3-6eca-f907-ae5b-4327c3c1618e
ms.date: 06/08/2017
---


# Close a Report Automatically If It Does Not Contain Any Records

The following example shows how to use the  **[NoData](report-nodata-event-access.md)** event to cancel opening or printing a report when it has no data. A message box notifying the user that the report has no data is also displayed.


```vb
Private Sub Report_NoData (Cancel As Integer) 
     
    ' Display message to user. 
    MsgBox "There are no records to report", vbExclamation, "No Records" 
 
    ' Close the report. 
    Cancel = True 
 
End Sub
```


