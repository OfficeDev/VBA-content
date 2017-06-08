---
title: Items.Restrict Method (Outlook)
keywords: vbaol11.chm70
f1_keywords:
- vbaol11.chm70
ms.prod: outlook
api_name:
- Outlook.Items.Restrict
ms.assetid: e3b0cda1-e43d-cc5e-2942-0f54935d9dab
ms.date: 06/08/2017
---


# Items.Restrict Method (Outlook)

Applies a filter to the  **[Items](items-object-outlook.md)** collection, returning a new collection containing all of the items from the original that match the filter.


## Syntax

 _expression_ . **Restrict**( **_Filter_** )

 _expression_ An expression that returns a **Items** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filter_|Required| **String**|A filter string expression to be applied. For details, see the  **[Find](items-find-method-outlook.md)** method.|

### Return Value

An  **Items** collection that represents the items from the original **Items** collection which match the filter.


## Remarks

This method is an alternative to using the  **[Find](items-find-method-outlook.md)** method or **[FindNext](items-findnext-method-outlook.md)** method to iterate over specific items within a collection. The **Find** or **FindNext** methods are faster than filtering if there are a small number of items. The **Restrict** method is significantly faster if there is a large number of items in the collection, especially if only a few items in a large collection are expected to be found.


 **Note**  If you are using user-defined fields as part of a  **Find** or **Restrict** clause, the user-defined fields must exist in the folder. Otherwise the code will generate an error stating that the field is unknown. You can add a field to a folder by displaying the **Field Chooser** and clicking **New**.

This method cannot be used and will cause an error with the following properties:



| **Body**| **LastFirstNoSpaceCompany**|
| **Categories**| **LastFirstSpaceOnly**|
| **Children**| **LastFirstSpaceOnlyCompany**|
| **Class**| **LastFirstNoSpaceAndSuffix**|
| **Companies**| **MemberCount**|
| **CompanyLastFirstNoSpace**| **NetMeetingAlias**|
| **CompanyLastFirstSpaceOnly**| **NetMeetingAutoStart**|
| **ContactNames**| **NetMeetingOrganizerAlias**|
| **Contacts**| **NetMeetingServer**|
| **ConversationIndex**| **NetMeetingType**|
| **DLName**| **RecurrenceState**|
| **Email1EntryID**| **ReceivedByEntryID**|
| **Email2EntryID**| **RecevedOnBehalfOfEntryID**|
| **Email3EntryID**| **ReplyRecipients**|
| **EntryID**| **ResponseState**|
| **HTMLBody**| **Saved**|
| **IsOnlineMeeting**| **Sent**|
| **LastFirstAndSuffix**| **Submitted**|
| **LastFirstNoSpace**| **VotingOptions**|
| **AutoResolvedWinner**| **DownloadState**|
| **BodyFormat**| **IsConflict**|
| **InternetCodePage**| **MeetingWorkspaceURL**|
| **Permission**||

### Creating Filters for the Find and Restrict Methods

The syntax for the filter varies depending on the type of field you are filtering on. 


### String (for Text fields)

When filtering text fields, you can use either a pair of single quotes (') or a pair of double quotes ("), to delimit the values that are part of the filter. For example, all of the following lines function correctly when the field is of type  **String** :


```
sFilter = "[CompanyName] = 'Microsoft'" sFilter = "[CompanyName] = ""Microsoft"""  
sFilter = "[CompanyName] = " &; Chr(34) &; "Microsoft" &; Chr(34)
```

In specifying a filter in a Jet or DASL query, if you use a pair of single quotes to delimit a string that is part of the filter, and the string contains another single quote or apostrophe, then add a single quote as an escape character before the single quote or apostrophe. Use a similar approach if you use a pair of double quotes to delimit a string. If the string contains a double quote, then add a double quote as an escape character before the double quote. 

For example, in the DASL filter string that filters for the  **Subject** property being equal to the word `can't`, the entire filter string is delimited by a pair of double quotes, and the embedded string  `can't` is delimited by a pair of single quotes. There are three characters that you need to escape in this filter string: the starting double quote and the ending double quote for the property reference of `http://schemas.microsoft.com/mapi/proptag/0x0037001f`, and the apostrophe in the value condition for the word  `can't`. Applying the appropriate escape characters, you can express the filter string as follows: 




```
filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" = 'can''t'"
```

Alternatively, you can use the  `chr(34)` function to represent the double quote (whose ASCII character value is 34) that is used as an escape character. Using the `chr(34)` substitution for a double-quote escape character, you can express the last example as follows:




```
filter = "@SQL= " &; Chr(34) &; "http://schemas.microsoft.com/mapi/proptag/0x0037001f" _  
    &; Chr(34) &; " = " &; "'can''t'"
```

Escaping single and double quote characters is also required for DASL queries with the  **ci_startswith** or **ci_phrasematch** operators. For example, the following query performs a phrase match query for `can't` in the message subject:




```
filter = "@SQL=" &; Chr(34) &; "http://schemas.microsoft.com/mapi/proptag/0x0037001E" _  
    &; Chr(34) &; " ci_phrasematch " &; "'can''t'"
```

Another example is a DASL filter string that filters for the  **Subject** property being equal to the words `the right stuff`, where the word  `stuff` is enclosed by double quotes. In this case, you must escape the enclosing double quotes as follows:




```
filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" = 'the right ""stuff""'"
```

A different set of escaping rules apply to a property reference for named properties that contain the space, single quote, double quote, or percent character. For more information, see [Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).


### Date

Although dates and times are typically stored with a  **Date** format, the **Find** and **Restrict** methods require that the date and time be converted to a string representation. To make sure that the date is formatted as Microsoft Outlook expects, use the **Format** function. The following example creates a filter to find all contacts that have been modified after January 15, 1999 at 3:30 P.M.


```
sFilter = "[LastModificationTime] > '" &; Format("1/15/99 3:30pm", "ddddd h:nn AMPM") &; "'"
```


### Boolean Operators

 **Boolean** operators, **TRUE**/ **FALSE**, YES/NO, ON/OFF, and so on, should not be converted to a string. For example, to determine whether journaling is enabled for contacts, you can use this filter: 


```
sFilter = "[Journal] = True" 
```


 **Note**  If you use quotation marks as delimiters with  **Boolean** fields, then an empty string will find items whose fields are **False** and all non-empty strings will find items whose fields are **True** .


### Keywords (or Categories)

The  **Categories** field is of type keywords, which is designed to hold multiple values. When accessing it programmatically, the **Categories** field behaves like a Text field, and the string must match exactly. Values in the text string are separated by a comma and a space. This typically means that you cannot use the **Find** and **Restrict** methods on a keywords field if it contains more than one value. For example, if you have one contact in the Business category and one contact in the Business and Social categories, you cannot easily use the **Find** and **Restrict** methods to retrieve all items that are in the Business category. Instead, you can loop through all contacts in the folder and use the **Instr** function to test whether the string "Business" is contained within the entire keywords field.


 **Note**  A possible exception is if you limit the Categories field to two, or a low number of values. Then you can use the  **Find** and **Restrict** methods with the OR logical operator to retrieve all Business contacts. For example (in pseudocode): "Business" OR "Business, Personal" OR "Personal, Business." Category strings are not case sensitive.


### Integer

You can search for  **Integer** fields with, or without quotation marks as delimiters. The following filters will find contacts that were created by using Outlook 2000:


```
sFilter = "[OutlookInternalVersion] = 92711" sFilter = "[OutlookInternalVersion] = '92711'"
```


### Using Variables as Part of the Filter

As the  **Restrict** method example illustrates, you can use values from variables as part of the filter. The following Microsoft Visual Basic Scripting Edition (VBScript) code sample illustrates syntax that uses variables as part of the filter.


```
sFullName = "Dan Wilson"  
' This approach uses Chr(34) to delimit the value.  
sFilter = "[FullName] = " &; Chr(34) &; sFullName &; Chr(34)  
' This approach uses double quotation marks to delimit the value. sFilter = "[FullName] = """ &; sFullName &; """"
```


### Using Logical Operators as Part of the Filter

Logical operators that are allowed are AND, OR, and NOT. The following are variations of the clause for the  **Restrict** method so you can specify multiple criteria.

OR: The following code returns all contact items that have either Business or Personal as their category. 




```vb
sFilter = "[Categories] = 'Personal' Or [Categories] = 'Business'"  

```

AND: The following code retrieves all personal contacts who work at Microsoft. 




```vb
sFilter = "[Categories] = 'Personal' And [CompanyName] = 'Microsoft'"
```

NOT: The following code retrieves all personal contacts who don't work at Microsoft. 




```vb
sFilter = "[Categories] = 'Personal' And Not([CompanyName] = 'Microsoft')"
```


### Additional Notes

If you are trying to use the  **Find** or **Restrict** methods with user-defined fields, the fields must be defined in the folder, otherwise an error will occur. There is no way to perform a "contains" operation. For example, you cannot use **Find** or **Restrict** to search for items that have a particular word in the **Subject** field. Instead, you can use the **[AdvancedSearch](application-advancedsearch-method-outlook.md)** method, or you can loop through all of the items in the folder and use the **InStr** function to perform a search within a field. You can use the **Restrict** method to search for items that begin within a certain range of characters. For example, to search for all contacts with a last name beginning with the letter M, use this filter:


```vb
sFilter = "[LastName] > 'LZZZ' And [LastName] < 'N'"
```


## Example

This Visual Basic for Applications (VBA) example uses the  **Restrict** method to get all Inbox items of **Business** category and moves them to the **Business** folder. To run this example, create or make sure a subfolder called 'Business' exists under Inbox.


```vb
Sub MoveItems()  
    Dim myNamespace As Outlook.NameSpace  
    Dim myFolder As Outlook.Folder  
    Dim myItems As Outlook.Items  
    Dim myRestrictItems As Outlook.Items  
    Dim myItem As Outlook.MailItem  
  
    Set myNamespace = Application.GetNamespace("MAPI")  
    Set myFolder = _  
        myNamespace.GetDefaultFolder(olFolderInbox)  
    Set myItems = myFolder.Items  
    Set myRestrictItems = myItems.Restrict("[Categories] = 'Business'")  
    For i =  myRestrictItems.Count To 1 Step -1  
        myRestrictItems(i).Move myFolder.Folders("Business")  
    Next  
End Sub
```

This Visual Basic for Applications example uses the  **Restrict** method to apply a filter to contact items based on the item's **[LastModificationTime](contactitem-lastmodificationtime-property-outlook.md)** property.




```vb
Public Sub ContactDateCheck()  
    Dim myNamespace As Outlook.NameSpace  
    Dim myContacts As Outlook.Items  
    Dim myItems As Outlook.Items  
    Dim myItem As Object  
      
    Set myNamespace = Application.GetNamespace("MAPI")  
    Set myContacts = myNamespace.GetDefaultFolder(olFolderContacts).Items  
    Set myItems = myContacts.Restrict("[LastModificationTime] > '01/1/2003'")  
    For Each myItem In myItems  
        If (myItem.Class = olContact) Then  
            MsgBox myItem.FullName &; ": " &; myItem.LastModificationTime  
        End If  
    Next  
End Sub
```

The following Visual Basic for Applications example is the same as the example above, except that it demonstrates the use of a variable in the filter.




```vb
Public Sub ContactDateCheck2()  
    Dim myNamespace As Outlook.NameSpace  
    Dim myContacts As Outlook.Items  
    Dim myItem As Object  
    Dim DateStart As Date  
    Dim DateToCheck As String  
    Dim myRestrictItems As Outlook.Items  
  
    Set myNameSpace = Application.GetNamespace("MAPI")  
    Set myContacts = myNameSpace.GetDefaultFolder(olFolderContacts).Items  
    DateStart = #01/1/2003#  
    DateToCheck = "[LastModificationTime] >= """ &; DateStart &; """"  
    Set myRestrictItems = myContacts.Restrict(DateToCheck)  
    For Each myItem In myRestrictItems  
        If (myItem.Class = olContact) Then  
            MsgBox myItem.FullName &; ": " &; myItem.LastModificationTime  
        End If  
    Next  
End Sub
```


## See also


#### Concepts


[Items Object](items-object-outlook.md)

