---
title: Filtering Items Using a Date-time Comparison
ms.prod: outlook
ms.assetid: 668e0993-c3d2-835f-0645-ba79bcffe67f
ms.date: 06/08/2017
---


# Filtering Items Using a Date-time Comparison

## Filtering Recurring Items in the Calendar Folder

To filter a collection of appointment items that include recurring appointments, you must use the  **[Items](items-object-outlook.md)** collection. Use the **[Items.IncludeRecurrences](items-includerecurrences-property-outlook.md)** property to specify that **[Items.Find](items-find-method-outlook.md)** or **[Items.Restrict](items-restrict-method-outlook.md)** should include recurring appointments. The **[Table](table-object-outlook.md)** object returns only one row representing the recurrent appointment item, instead of a row for each occurrence of the appointment.


## Date-time Format of Comparison Strings

Outlook evaluates date-time values according to the time format, short date format, and long date format settings in the Regional and Language Options applet in the Windows Control Panel. In particular, Outlook evaluates time according to that specified time format without seconds. If you specify seconds in the date-time comparison string, the filter will not operate as expected.

Although dates and times are typically stored with a date format, filters using the Jet and DAV Searching and Locating (DASL) syntax require that the date-time value to be converted to a string representation. In Jet syntax, the date-time comparison string should be enclosed in either double quotes or single quotes. In DASL syntax, the date-time comparison string should be enclosed in single quotes.

To make sure that the date-time comparison string is formatted as Microsoft Outlook expects, use the Visual Basic for Applications  **Format** function (or its equivalent in your programming language). The following example creates a Jet filter to find all contacts that have been modified before June 12, 2005 at 3:30 P.M local time.




```
criteria = "[LastModificationTime] < '" _ 
         &; Format$("6/12/2005 3:30PM","General Date") &; "'"
```


## Time Zones Used in Comparison

When an explicit built-in property is referenced in a Jet query with its explict string name, the comparison evaluates the property value and the date-time comparison string as local time values.

When a property is referenced in a DASL query by namespace, the comparison evaluates the property value and the date-time comparison string as Coordinated Universal Time (UTC) values. For example, the following DASL query finds all contacts that have been modified before June 12, 2005 at 3:30 pm, UTC.




```
criteria = "@SQL=" &; Chr(34) &; "DAV:getlastmodified" &; Chr(34) _ 
         &; " < '" &; Format$("6/12/2005 3:30PM","General Date") &; "'"
```


## Conversion to UTC for DASL Queries

Since DASL queries always perform date-time comparisons in UTC, if you use a date literal in a comparison string, you must use its UTC value for the comparison. You can use the  **[Row.LocalTimeToUTC](row-localtimetoutc-method-outlook.md)** helper function or Outlook date-time macros to facilitate the conversion.


## LocalTimeToUTC

 One way to facilitate local time to UTC conversion is to use the helper function, **LocalTimeToUTC**, of the  **[Row](row-object-outlook.md)** object. The following line of code uses this helper function to convert the value of the **LastModificationTime** property (which is a default column in all **Table** objects):


```
Row.LocalTimeToUTC("LastModificationTime")
```


## Outlook Date-time Macros

The date macros listed below return filter strings that compare the value of a given date-time property with a specified date in UTC;  _SchemaName_ is any valid date-time property referenced by namespace.


 **Note**  Outlook date-time macros can be used only in DASL queries.



| **Macro**| **Syntax**| **Description**|
|:-----|:-----|:-----|
|today|%today(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to today|
|tomorrow|%tomorrow(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to tomorrow|
|yesterday|%yesterday(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to yesterday|
|next7days|%next7days(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to next 7 days|
|last7days|%last7days(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to last 7 days|
|nextweek|%nextweek(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to next week|
|thisweek|%thisweek(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to this week|
|lastweek|%lastweek(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to last week|
|nextmonth|%nextmonth(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to next month|
|thismonth|%thismonth(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to this month|
|lastmonth|%lastmonth(" _SchemaName_")%|Restricts for items with  _SchemaName_ property value equal to last month|

## Example Showing Conversion to UTC

The following code example illustrates three filter strings that return all messages received today, and applies one of the filters to  **Items.Restrict** and **[Application.AdvancedSearch](application-advancedsearch-method-outlook.md)**. It first uses  **[PropertyAccessor.LocalTimeToUTC](propertyaccessor-localtimetoutc-method-outlook.md)** to convert today's date to UTC date strings. The first filter uses the Outlook macro, **today**, to obtain a filter string that compares the  **ReceivedTime** property with today's date in UTC. The second and third macros reference the **ReceivedTime** property by two different namespaces. 

The code example finally applies the third filter to items in the Inbox twice, first using **Items.Restrict** and then using **Application.AdvancedSearch**. It prints the number of items in the Inbox, and the number of items returned from each application of the filter.


```vb
Public blnSearchComp As Boolean 
 
Sub TestDASLDateComparison() 
    Dim strFilter As String 
    Dim colItems As Outlook.Items 
    Dim colRestrict As Outlook.Items 
    Dim oSearch As Outlook.Search 
    Dim oResults As Outlook.Results 
    Dim datStartUTC As Date 
    Dim datEndUTC As Date 
    Dim oMail As MailItem 
    Dim oPA As PropertyAccessor 
    Const SchemaPropTag As String = _ 
    "http://schemas.microsoft.com/mapi/proptag/" 
 
    'Get items from Inbox 
    Set colItems = _ 
    Application.Session.GetDefaultFolder(olFolderInbox).Items 
     
    'This code is a workaround to get today's date 
    'as UTC for DASL date comparison 
    Set oMail = Application.CreateItem(olMailItem) 
    Set oPA = oMail.PropertyAccessor 
    datStartUTC = oPA.LocalTimeToUTC(Date) 
    datEndUTC = oPA.LocalTimeToUTC(DateAdd("d", 1, Date)) 
     
    'All three filters shown below will return the same results 
    'This filter uses DASL date macro for today 
    strFilter = "%today(" _ 
    &; AddQuotes("urn:schemas:httpmail:datereceived") &; ")%" 
     
    'This filter uses urn:schemas:httpmail namespace 
    strFilter = AddQuotes("urn:schemas:httpmail:datereceived") _ 
    &; " > '" &; datStartUTC &; "' AND " _ 
    &; AddQuotes("urn:schemas:httpmail:datereceived") _ 
    &; " < '" &; datEndUTC &; "'" 
 
    'This filter uses http://schemas.microsoft.com/mapi/proptag 
    strFilter = AddQuotes(SchemaPropTag &; "0x0E060040") _ 
    &; " > '" &; datStartUTC &; "' AND " _ 
    &; AddQuotes(SchemaPropTag &; "0x0E060040") _ 
    &; " < '" &; datEndUTC &; "'" 
 
    'Count of items in Inbox 
    Debug.Print (colItems.Count) 
 
    'This call succeeds with @SQL prefix 
    Set colRestrict = colItems.Restrict("@SQL=" &; strFilter) 
    'Get count of restricted items 
    Debug.Print (colRestrict.Count) 
 
    Set oSearch = Application.AdvancedSearch("Inbox", strFilter, False) 
    While blnSearchComp = False 
        DoEvents 
    Wend      
 
    'Get count from Search object 
    Set oResults = oSearch.Results 
    Debug.Print (oResults.Count) 
End Sub 
 
Public Function AddQuotes(ByVal SchemaName As String) As String 
    On Error Resume Next 
    AddQuotes = Chr(34) &; SchemaName &; Chr(34) 
End Function 
 
 
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
    MsgBox "The AdvancedSearchComplete Event fired" 
    blnSearchComp = True 
End Sub
```


