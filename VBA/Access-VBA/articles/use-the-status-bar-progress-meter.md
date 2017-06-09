---
title: Use the Status Bar Progress Meter
ms.prod: access
ms.assetid: 1ced64d3-56e4-064e-3dd2-d6b5e4dbdd8a
ms.date: 06/08/2017
---


# Use the Status Bar Progress Meter

This topic shows how to use the  **[SysCmd](application-syscmd-method-access.md)** method to create a progress meter on the status bar that gives a visual representation of the progress of an operation that has a known duration or number of steps.

There are three intrinsic constants that can be used with the  **SysCmd** method's _action_ argument to manipulate the progress meter on the status bar. The following table describes them.


|**Intrinsic constant**|**Description**|
|:-----|:-----|
|**acSysCmdInitMeter**|Initialize the progress meter. The maximum value that the process will attain is specifed in the  **SysCmd** method's _value_ argument.|
|**acSysCmdUpdateMeter**|Update the progress meter. A numeric expression that represents the current progress toward completion is specified in the  **SysCmd** method's _value_ argument.|
|**acSysCmdRemoveMeter**|Remove progress meter.|
The following procedure uses the  **SysCmd** method to update the progress meter as data from the Customers table is printed in the Immediate window.



```vb
Sub ProgressMeter() 
   Dim MyDB As DAO.Database, MyTable As DAO.Recordset 
   Dim Count As Long 
   Dim Progress_Amount As Integer 
    
   Set MyDB = CurrentDb() 
   Set MyTable = MyDB.OpenRecordset("Customers") 
 
   ' Move to last record of the table to get the total number of records. 
   MyTable.MoveLast 
   Count = MyTable.RecordCount 
 
   ' Move back to first record. 
   MyTable.MoveFirst 
 
   ' Initialize the progress meter. 
    SysCmd acSysCmdInitMeter, "Reading Data...", Count 
 
   ' Enumerate through all the records. 
   For Progress_Amount = 1 To Count 
     ' Update the progress meter. 
      SysCmd acSysCmdUpdateMeter, Progress_Amount 
       
     'Print the contact name and number of orders in the Immediate window. 
      Debug.Print MyTable![ContactName]; _ 
                  DCount("[OrderID]", "Orders", "[CustomerID]='" &; MyTable![CustomerID] &; "'") 
                   
     ' Go to the next record. 
      MyTable.MoveNext 
   Next Progress_Amount 
 
   ' Remove the progress meter. 
   SysCmd acSysCmdRemoveMeter 
         
End Sub
```


