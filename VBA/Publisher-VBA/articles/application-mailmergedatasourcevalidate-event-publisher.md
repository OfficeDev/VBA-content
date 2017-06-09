---
title: Application.MailMergeDataSourceValidate Event (Publisher)
keywords: vbapb10.chm268435480
f1_keywords:
- vbapb10.chm268435480
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeDataSourceValidate
ms.assetid: 8e18b0a0-8fe8-f72e-8a75-1585367cc796
ms.date: 06/08/2017
---


# Application.MailMergeDataSourceValidate Event (Publisher)

Occurs when a user performs address verification by clicking  **Validate** in the **Mail Merge Recipients** dialog box.


## Syntax

 _expression_. **MailMergeDataSourceValidate**( **_Doc_**,  **_Handled_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The mail merge main document.|
|Handled|Required| **Boolean**| **True** runs the accompanying validation code against the mail merge data source. **False** cancels the data source validation.|

## Remarks

If you do not have address verification software installed on your computer, use the  **MailMergeDataSourceValidate** event to create simple filtering routines, such as looping through records to check the postal codes and remove any that are non-U.S. Non-U.S. users can filter out all U.S. postal codes by modifying the code sample below and using Microsoft Visual Basic commands to search for text or special characters.

To access the  **Application** object events, declare an **Application** object variable in the General Declarations section of a code module. Then set the variable equal to the **Application** object for which you want to access events. For information about using events with the Microsoft Publisher **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md).


## Example

This example validates ZIP Codes in the attached data source for five digits. If the length of the ZIP Code is fewer than five digits, the record is excluded from the mail merge process. This example assumes the postal codes are U.S. ZIP Codes. You could modify this example to search for ZIP Codes that have a four-digit locator code appended to the ZIP Code, and then exclude all records that do not contain the locator code.


```vb
Private Sub MailMergeApp_MailMergeDataSourceValidate( _ 
 ByVal Doc As Document, _ 
 Handled As Boolean) 
 
 Dim intCount As Integer 
 
 Handled = True 
 
 On Error Resume Next 
 
 With ActiveDocument.MailMerge.DataSource 
 
 'Set the active record equal to the first included record in the 
 'data source 
 .ActiveRecord = 1 
 Do 
 intCount = intCount + 1 
 
 'Set the condition that field six must be greater than or 
 'equal to five 
 If Len(.DataFields.Item(6).Value) < 5 Then 
 
 'Exclude the record if field six is shorter than five digits 
 .Included = False 
 
 'Mark the record as containing an invalid address field 
 .InvalidAddress = True 
 
 'Specify the comment attached to the record explaining 
 'why the record was excluded from the mail merge 
 .InvalidComments = "The ZIP Code for this record has " _ 
 &; "fewer than five digits. It will be removed " _ 
 &; "from the mail merge process." 
 
 End If 
 
 'Move the record to the next record in the data source 
 .ActiveRecord = .ActiveRecord + 1 
 
 'End the loop when the counter variable 
 'equals the number of records in the data source 
 Loop Until intCount = .RecordCount 
 End With 
 
End Sub
```

For this event to occur, you must place the following line of code in the General Declarations section of your module and run the following initialization routine.




```vb
Private WithEvents MailMergeApp As Application 
 
Sub InitializeMailMergeApp() 
 Set MailMergeApp = Publisher.Application 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

