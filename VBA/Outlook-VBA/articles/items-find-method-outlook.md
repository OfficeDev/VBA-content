---
title: Items.Find Method (Outlook)
keywords: vbaol11.chm62
f1_keywords:
- vbaol11.chm62
ms.prod: outlook
api_name:
- Outlook.Items.Find
ms.assetid: e7a791d8-b80b-df07-84a3-a85acabfcf80
ms.date: 06/08/2017
---


# Items.Find Method (Outlook)

Locates and returns a Microsoft Outlook item object that satisfies the given  _Filter_ .


## Syntax

 _expression_ . **Find**( **_Filter_** )

 _expression_ An expression that returns a **Items** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filter_|Required| **String**|A string that specifies the criteria that the returned object must satisfy.|

### Return Value

An  **Object** value that represents an Outlook item if the call succeeds; returns **Null** (or **Nothing** in Visual Basic) if it fails.


## Remarks

To use content indexing search in the  **[Items](items-object-outlook.md)** collection, use the **[Restrict](items-restrict-method-outlook.md)** method. **FindRow** will return an error if _Filter_ contains content indexing keywords. For more information on content indexing keywords, see[Filtering Items Using Query Keywords](http://msdn.microsoft.com/library/d7e6b169-c5fd-7acc-f077-658a153a921f%28Office.15%29.aspx).

The method will return an error with the following properties in the  _Filter_ :



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
| **Email1EntryID**| **ReplyRecipients**|
| **Email2EntryID**| **ReceivedByEntryID**|
| **Email3EntryID**| **RecevedOnBehalfOfEntryID**|
| **EntryID**| **ResponseState**|
| **HTMLBody**| **Saved**|
| **IsOnlineMeeting**| **Sent**|
| **LastFirstAndSuffix**| **Submitted**|
| **LastFirstNoSpace**| **VotingOptions**|
| **AutoResolvedWinner**| **DownloadState**|
| **BodyFormat**| **IsConflict**|
| **InternetCodePage**| **MeetingWorkspaceURL**|
| **Permission**||
 **Creating Filters for the Find and Restrict Methods**

The syntax for the filter varies depending on the type of field you are filtering on.

 **String (for Text fields)**

When filtering text fields, you can use either a pair of single quotes (') or a pair of double quotes (") to delimit the values that are part of the filter. For example, all of the following lines function correctly when the field is of type  **String** :




```
sFilter = "[CompanyName] = 'Microsoft'"  
sFilter = "[CompanyName] = ""Microsoft"""  
sFilter = "[CompanyName] = " &; Chr(34) &; "Microsoft" &; Chr(34)
```

In specifying a filter in a Jet or DASL query, if you use a pair of single quotes to delimit a string that is part of the filter, and the string contains another single quote or apostrophe, then add a single quote as an escape character before the single quote or apostrophe. Use a similar approach if you use a pair of double quotes to delimit a string. If the string contains a double quote, then add a double quote as an escape character before the double quote. 

For example, in the DASL filter string that filters for the  **Subject** property being equal to the word `can't`, the entire filter string is delimited by a pair of double quotes, and the embedded string  `can't` is delimited by a pair of single quotes. There are three characters that you need to escape in this filter string: the starting double quote and the ending double quote for the property reference of `http://schemas.microsoft.com/mapi/proptag/0x0037001f`, and the apostrophe in the value condition for the word  `can't`. Applying the appropriate escape characters, you can express the filter string as follows: 




```
filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" = 'can''t'"
```

Alternatively, you can use the  `chr(34)` function to represent the double quote (whose ASCII character value is 34) that is used as an escape character. Using the `chr(34)` substitution for a double-quote escape character, you can express the last example as follows:




```
filter = "@SQL= " &; Chr(34) &; "http://schemas.microsoft.com/mapi/proptag/0x0037001f" _&; Chr(34) &; " = " &; "'can''t'"
```

Escaping single and double quote characters is also required for DASL queries with the  **ci_startswith** or **ci_phrasematch** operators. For example, the following query performs a phrase match query for `can't` in the message subject:




```
filter = "@SQL=" &; Chr(34) &; "http://schemas.microsoft.com/mapi/proptag/0x0037001E" _&; Chr(34) &; " ci_phrasematch " &; "'can''t'"
```

Another example is a DASL filter string that filters for the  **Subject** property being equal to the words `the right stuff`, where the word  `stuff` is enclosed by double quotes. In this case, you must escape the enclosing double quotes as follows:




```
filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" = 'the right ""stuff""'"
```

A different set of escaping rules apply to a property reference for named properties that contain the space, single quote, double quote, or percent character. For more information, see [Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).

 **Date**

Although dates and times are typically stored with a  **Date** format, the **Find** and **Restrict** methods require that the date and time be converted to a string representation. To make sure that the date is formatted as Outlook expects, use the **Format** function. The following example creates a filter to find all contacts that have been modified after January 15, 1999 at 3:30 P.M.




```
sFilter = "[LastModificationTime] > '" &; Format("1/15/99 3:30pm", "ddddd h:nn AMPM") &; "'"
```

 **Boolean Operators**

 **Boolean** operators, **TRUE**/ **FALSE**, YES/NO, ON/OFF, and so on, should not be converted to a string. For example, to determine whether journaling is enabled for contacts, you can use this filter: 




```
sFilter = "[Journal] = True" 
```


 **Note**  If you use quotation marks as delimiters with  **Boolean** fields, then an empty string will find items whose fields are **False** and all non-empty strings will find items whose fields are **True**.

 **Keywords (or Categories)**

The  **Categories** field is of type keywords, which is designed to hold multiple values. When accessing it programmatically, the **Categories** field behaves like a Text field, and the string must match exactly. Values in the text string are separated by a comma and a space. This typically means that you cannot use the **Find** and **Restrict** methods on a keywords field if it contains more than one value. For example, if you have one contact in the Business category and one contact in the Business and Social categories, you cannot easily use the **Find** and **Restrict** methods to retrieve all items that are in the Business category. Instead, you can loop through all contacts in the folder and use the **Instr** function to test whether the string "Business" is contained within the entire keywords field.


 **Note**  A possible exception is if you limit the  **Categories** field to two, or a low number of values. Then you can use the **Find** and **Restrict** methods with the OR logical operator to retrieve all Business contacts. For example (in pseudocode): "Business" OR "Business, Personal" OR "Personal, Business." Category strings are not case sensitive.

 **Integer**

You can search for  **Integer** fields with or without quotation marks as delimiters. The following filters will find contacts that were created with Outlook 2000:




```
sFilter = "[OutlookInternalVersion] = 92711"  
sFilter = "[OutlookInternalVersion] = '92711'"
```

 **Using Variables as Part of the Filter**

As the  **Restrict** method example illustrates, you can use values from variables as part of the filter. The following Microsoft Visual Basic Scripting Edition (VBScript) code sample illustrates syntax that uses variables as part of the filter.




```
sFullName = "Dan Wilson" 
```

 This approach uses Chr(34) to delimit the value:




```
sFilter = "[FullName] = " &; Chr(34) &; sFullName &; Chr(34)
```

 This approach uses double quotation marks to delimit the value:




```
sFilter = "[FullName] = """ &; sFullName &; """"
```

 **Using Logical Operators as Part of the Filter**

Logical operators that are allowed are AND, OR, and NOT. The following are variations of the clause for the  **Restrict** method, so you can specify multiple criteria.

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

 **Additional Notes**

If you are trying to use the  **Find** or **Restrict** methods with user-defined fields, the fields must be defined in the folder, otherwise an error will occur. There is no way to perform a "contains" operation. For example, you cannot use **Find** or **Restrict** to search for items that have a particular word in the **Subject** field. Instead, you can use the **AdvancedSearch** method, or you can loop through all of the items in the folder and use the **InStr** function to perform a search within a field. You can use the **Restrict** method to search for items that begin within a certain range of characters. For example, to search for all contacts with a last name beginning with the letter M, use this filter:




```vb
sFilter = "[LastName] > 'LZZZ' And [LastName] < 'N'"
```


## See also


#### Concepts


[Items Object](items-object-outlook.md)

