---
title: Work With Attachments In DAO
ms.prod: access
ms.assetid: e175a47a-4d97-b93b-c152-809314ac5ba0
ms.date: 06/08/2017
---


# Work With Attachments In DAO

In DAO, Attachment fields function just like other multi-valued fields. The field that contains the attachment contains a recordset that is a child to the table's recordset. There are two new DAO methods,  **[LoadFromFile](http://msdn.microsoft.com/library/33FD543F-BD24-9199-7540-2889B69221C8%28Office.15%29.aspx)** and **[SaveToFile](http://msdn.microsoft.com/library/250F9596-1A03-471D-96F9-718CD57DC94F%28Office.15%29.aspx)**, that deal exclusively with attachments.


## Add an Attachment to a Record

The  **LoadFromFile** method loads a file from disk and adds the file as an attachment to the specified record. The following code example shows the syntax of the **LoadFromFile** method.


```vb
Recordset.Fields("FileData").LoadFromFile(<filename>)
```


 **Note**  The  **FileData** field is reserved internally by the Access database engine to store the binary attachment data.

The following code example uses the  **LoadFromFile** method to load an employee's picture from disk.




```vb
   '  Instantiate the parent recordset.  
   Set rsEmployees = db.OpenRecordset("Employees") 
  
   … Code to move to desired employee 
  
   ' Activate edit mode. 
   rsEmployees.Edit 
  
   ' Instantiate the child recordset. 
   Set rsPictures = rsEmployees.Fields("Pictures").Value  
  
   ' Add a new attachment. 
   rsPictures.AddNew 
   rsPictures.Fields("FileData").LoadFromFile "EmpPhoto39392.jpg" 
   rsPictures.Update 
  
   ' Update the parent record 
   rsEmployees.Update 

```


## Save an Attachment to Disk

The following code example shows how to use the  **SaveToFile** method to save all of the attachments for a specific employee to disk.


```vb
'  Instantiate the parent recordset.  
   Set rsEmployees = db.OpenRecordset("Employees") 
  
   … Code to move to desired employee 
  
   ' Instantiate the child recordset. 
   Set rsPictures = rsEmployees.Fields("Pictures").Value  
 
   '  Loop through the attachments. 
   While Not rsPictures.EOF 
  
      '  Save current attachment to disk in the "My Documents" folder. 
      rsPictures.Fields("FileData").SaveToFile _ 
                  "C:\Documents and Settings\Username\My Documents" 
      rsPictures.MoveNext 
   Wend 

```


