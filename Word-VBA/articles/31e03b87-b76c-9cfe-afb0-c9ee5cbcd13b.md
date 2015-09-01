
# Application.MailMergeDataSourceValidate Event (Word)

 **Last modified:** July 28, 2015

Occurs when a user validates mail merge recipients by clicking  **Validate** in the **Mail Merge Recipients** dialog box.

## Syntax

Private Sub  _expression__**MailMergeDataSourceValidate**( **_ByVal Doc As Document_**,  **_Handled As Boolean_**)

 _expression_A variable that represents an  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object that has been declared with events in a class module.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The mail merge main document.|
|Handled|Required| **Boolean**| **True** if the add-in has handled the validation event. This is a forward-only parameter and cannot be set in code. To set this value, you must use the ** [MailMergeDataSourceValidate2](dba0dc60-a8c7-7e0c-ac02-4f5311534c89.md)** event.|

## Remarks

If you do not have address verification software installed on your computer, the  **MailMergeDataSourceValidate** event allows you to create simple filtering routines, such as looping through records to check the postal codes and removing any that are non-U.S.


 **Note**  The Handled parameter does not function correctly in this version of the event; use the  ** [MailMergeDataSourceValidate2](dba0dc60-a8c7-7e0c-ac02-4f5311534c89.md)** event. In addition, you cannot raise this event from within a Microsoft Visual Basic for Applications (VBA) project. This event functions correctly only in COM add-ins. For managed add-ins and external applications, use the **MailMergeDataSourceValidate2** event.

For information about using events with the  **Application** object, see [Using Events with the Application Object](784c4c61-7e47-3dbf-46f6-da655f786ca1.md).


## See also


#### Concepts


 [Application Object](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)
#### Other resources


 [Application Object Members](71669f1e-65f1-b0f1-b67d-355dfdbebe50.md)
