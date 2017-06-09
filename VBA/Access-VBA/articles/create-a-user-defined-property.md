---
title: Create a User-Defined Property
ms.prod: access
ms.assetid: 49d2fede-2fb5-0b1b-42cd-6147756ea1ca
ms.date: 06/08/2017
---


# Create a User-Defined Property

The following example attempts to set the value of a user-defined property. If the property does not exist, it uses the  **[CreateProperty](http://msdn.microsoft.com/library/F2039BE9-5FD8-F673-DFBF-0A71540CDC98%28Office.15%29.aspx)** method to create and set the value of the new property.


```vb
Sub CreatePropertyX() 
 
   Dim dbsNorthwind As Database 
   Dim prpLoop As Property 
 
   Set dbsNorthwind = OpenDatabase("Northwind.mdb") 
 
   ' Set the Archive property to True. 
   SetProperty dbsNorthwind, "Archive", True 
    
   With dbsNorthwind 
      Debug.Print "Properties of " &; .Name 
       
      ' Enumerate Properties collection of the Northwind  
      ' database. 
      For Each prpLoop In .Properties 
         If prpLoop <> "" Then Debug.Print "  " &; _ 
            prpLoop.Name &; " = " &; prpLoop 
      Next prpLoop 
 
      ' Delete the new property because this is a  
      ' demonstration. 
      .Properties.Delete "Archive" 
 
      .Close 
   End With 
 
End Sub 
 
Sub SetProperty(dbsTemp As Database, strName As String, _ 
   booTemp As Boolean) 
 
   Dim prpNew As Property 
   Dim errLoop As Error 
 
   ' Attempt to set the specified property. 
   On Error GoTo Err_Property 
   dbsTemp.Properties("strName") = booTemp 
   On Error GoTo 0 
 
   Exit Sub 
 
Err_Property: 
 
   ' Error 3270 means that the property was not found. 
   If DBEngine.Errors(0).Number = 3270 Then 
      ' Create property, set its value, and append it to the  
      ' Properties collection. 
      Set prpNew = dbsTemp.CreateProperty(strName, _ 
         dbBoolean, booTemp) 
      dbsTemp.Properties.Append prpNew 
      Resume Next 
   Else 
      ' If different error has occurred, display message. 
      For Each errLoop In DBEngine.Errors 
         MsgBox "Error number: " &; errLoop.Number &; vbCr &; _ 
            errLoop.Description 
      Next errLoop 
      End 
   End If 
 
End Sub
```


