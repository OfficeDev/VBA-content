---
title: Set Options from Visual Basic
ms.prod: access
ms.assetid: c85ab081-6522-f851-a0d7-3d6612af26ab
ms.date: 06/08/2017
---


# Set Options from Visual Basic

You can use the  **[SetOption](application-setoption-method-access.md)** and **[GetOption](application-getoption-method-access.md)** methods to set and return option values in the **Access Options** dialog box from code. To view the **Access Options** dialog box, click the Microsoft Office Button and then click **Access Options**.

The value that you pass to the  **SetOption** method as the _setting_ argument depends on which type of option you are setting. The following table establishes some guidelines for setting options.


|**If the option is**|**Then the  _setting_ argument is**|
|:-----|:-----|
|A text box|A string|
|A check box|A Boolean value —  **True** (-1) or **False** (0)|
|An option button in an option group, or an option in a combo box or a list box|An integer corresponding to the option's position in the option group or list (starting with zero [0])|

The following tables list the names of all options that can be set or returned from code and the tabs on which they can be found in the  **Access Options** dialog box, followed by the corresponding string argument that you must pass to the **SetOption** or **GetOption** method.


## Popular Tab
**Creating Databases Section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**New database sort order**|New Database Sort Order|
|**Default database folder**|Default Database Directory|
|**Default file format**|Default File Format|

## Current Database Tab

**Application Options section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Compact on Close**|Auto Compact|
|**Remove personal information from file properties on save**|Remove Personal Information|
|**Use Windows-themed Controls on Forms**|Themed Form Controls|
|**Enable Layout View for this database**|DesignWithData|
|**Check for truncated number fields**|CheckTruncatedNumFields|
|**Picture Property Storage Format**|Picture Property Storage Format|

**Name AutoCorrect Options section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Track name AutoCorrect info**|Track Name AutoCorrect Info|
|**Perform name AutoCorrect**|Perform Name AutoCorrect|
|**Log name AutoCorrect changes**|Log Name AutoCorrect Changes|

**Filter Lookup options for <Database Name> Database section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Show list of values in, Local indexed fields**|Show Values in Indexed|
|**Show list of values in, Local nonindexed fields**|Show Values in Non-Indexed|
|**Show list of values in, ODBC fields**|Show Values in Remote|
|**Show list of values in, Records in local snapshot**|Show Values in Snapshot|
|**Show list of values in, Records at server**|Show Values in Server|
|**Don't display lists where more than this number of records read**|Show Values Limit|

## Datasheet Tab

**Default colors section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Font color**|Default Font Color|
|**Background color**|Default Background Color|
|**Alternate background color**|_64|
|**Gridlines color**|Default Gridlines Color|

**Gridlines and cell effects section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Default gridlines showing, Horizontal**|Default Gridlines Horizontal|
|**Default gridlines showing, Vertical**|Default Gridlines Vertical|
|**Default cell effect**|Default Cell Effect|
|**Default column width**|Default Column Width|

 **Default font section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Font**|Default Font Name|
|**Size**|Default Font Size|
|**Weight**|Default Font Weight|
|**Underline**|Default Font Underline|
|**Italic**|Default Font Italic|

## Object Designers Tab

**Table design section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Default text field size**|Default Text Field Size|
|**Default number field size**|Default Number Field Size|
|**Default field type**|Default Field Type|
|**AutoIndex on Import/Create**|AutoIndex on Import/Create|
|**Show Property Update Option Buttons**|Show Property Update Options Buttons|

**Query design section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Show table names**|Show Table Names|
|**Output all fields**|Output All Fields|
|**Enable AutoJoin**|Enable AutoJoin|
|**SQL Server Compatible Syntax (ANSI 92), This database**|ANSI Query Mode|
|**SQL Server Compatible Syntax (ANSI 92), Default for new databases**|ANSI Query Mode Default|
|**Query design font, Font**|Query Design Font Name|
|**Query design font, Size**|Query Design Font Size|

**Forms/Reports section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Selection behavior**|Selection Behavior|
|**Form template**|Form Template|
|**Report template**|Report Template|
|**Always use event procedures**|Always Use Event Procedures|

**Error checking section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Enable error checking**|Enable Error Checking|
|**Error indicator color**|Error Checking Indicator Color|
|**Check for unassociated label and control**|Unassociated Label and Control Error Checking|
|**Check for new unassociated labels**|New Unassociated Labels Error Checking|
|**Check for keyboard shortcut errors**|Keyboard Shortcut Errors Error Checking|
|**Check for invalid control properties**|Invalid Control Properties Error Checking|
|**Check for common report errors**|Common Report Errors Error Checking|

## Proofing Tab

**When correcting spelling in Microsoft Office programs section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Ignore words in UPPERCASE**|Spelling ignore words in UPPERCASE|
|**Ignore words that contain numbers**|Spelling ignore words with number|
|**Ignore Internet and file addresses**|Spelling ignore Internet and file addresses|
|**Suggest from main dictionary only**|Spelling suggest from main dictionary only|
|**Dictionary Language**|Spelling dictionary language|

## Advanced Tab

**Editing section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Move after enter**|Move After Enter|
|**Behavior entering field**|Behavior Entering Field|
|**Arrow key behavior**|Arrow Key Behavior|
|**Cursor stops at first/last field**|Cursor Stops at First/Last Field|
|**Default find/replace behavior**|Default Find/Replace Behavior|
|**Confirm, Record changes**|Confirm Record Changes|
|**Confirm, Document deletions**|Confirm Document Deletions|
|**Confirm, Action queries**|Confirm Action Queries|
|**Default direction**|Default Direction|
|**General alignment**|General Alignment|
|**Cursor movement**|Cursor Movement|
|**Datasheet IME control**|Datasheet Ime Control|
|**Use Hijri Calendar**|Use Hijri Calendar|

**Display section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Show this number of Recent Documents**|Size of MRU File List|
|**Status bar**|Show Status Bar|
|**Show animations**|Show Animations|
|**Show Smart Tags on Datasheets**|Show Smart Tags on Datasheets|
|**Show Smart Tags on Forms and Reports**|Show Smart Tags on Forms and Reports|
|**Show in Macro Design, Names column**|Show Macro Names Column|
|**Show in Macro Design, Conditions column**|Show Conditions Column|

**Printing section**


|**Option Text**|**String Argument**|
|:-----|:-----|
|**Left margin**|Left Margin|
|**Right margin**|Right Margin|
|**Top margin**|Top Margin|
|**Bottom margin**|Bottom Margin|

**General section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Provide feedback with sound**|Provide Feedback with Sound|
|**Use four-year digit year formatting, This database**|Four-Digit Year Formatting|
|**Use four-year digit year formatting, All databases**|Four-Digit Year Formatting All Databases|

**Advanced Section**

|**Option Text**|**String Argument**|
|:-----|:-----|
|**Open last used database when Access starts**|Open Last Used Database When Access Starts|
|**Default open mode**|Default Open Mode for Databases|
|**Default record locking**|Default Record Locking|
|**Open databases by using record-level locking**|Use Row Level Locking|
|**OLE/DDE timeout (sec)**|OLE/DDE Timeout (sec)|
|**Refresh interval (sec)**|Refresh Interval (sec)|
|**Number of update retries**|Number of Update Retries|
|**ODBC refresh interval (sec)**|ODBC Refresh Interval (sec)|
|**Update retry interval (msec)**|Update Retry Interval (msec)|
|**DDE operations, Ignore DDE requests**|Ignore DDE Requests|
|**DDE operations, Enable DDE refresh**|Enable DDE Refresh|
|**Command-line arguments**|Command-Line Arguments|


|**Note**|
|:-----|    
|<ul><li>If your database may run on a version of Access for a language other than the one in which you created it, then you must supply the arguments for the GetOption and SetOption methods in English.</li><li>Some options are available only within an Access database or Access project (.adp).</li><li>If you are developing a database application, add-in, library database, or referenced database, make sure that the Error Trapping option is set to 2 (Break On Unhandled Errors) when you have finished debugging your code.</li></ul>|

