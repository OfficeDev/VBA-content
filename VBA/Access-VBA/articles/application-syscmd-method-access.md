---
title: Application.SysCmd Method (Access)
keywords: vbaac10.chm12515
f1_keywords:
- vbaac10.chm12515
ms.prod: access
api_name:
- Access.Application.SysCmd
ms.assetid: 5064b8cc-6f9a-602b-e304-6d1478d9b4a7
ms.date: 06/08/2017
---


# Application.SysCmd Method (Access)

You can use the  **SysCmd** method to, display a progress meter or optional specified text in the status bar, return information about Microsoft Access and its associated files, or return the state of a specified database object (to indicate whether the object is open, is a new object, or has been changed but not saved). **Variant**.


## Syntax

 _expression_. **SysCmd**( ** _Action_**, ** _Argument2_**, ** _Argument3_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Action_|Required|**AcSysCmdAction**|A  **[AcSysCmdAction](acsyscmdaction-enumeration-access.md)** constant that identifies the type of action to take. This set of constants applies to a progress meter. The **SysCmd** method returns a **null** if these actions are successful. Otherwise, Microsoft Access generates a run-time error.|
| _Argument2_|Optional|**Variant**|The text to be displayed left-aligned in the status bar. This argument is required when the action argument is  **acSysCmdInitMeter**, **acSysCmdUpdateMeter**, or **acSysCmdSetStatus**; this argument isn't valid for other action argument values.
 **Note**  When you specify the  **acSysCmdGetObjectState** value for the Action parameter, then you must specify the appropriate **[AcObjectType](acobjecttype-enumeration-access.md)** constant.

|
| _Argument3_|Optional|**Variant**|A numeric expression that controls the display of the progress meter. This argument is required when the action argument is  **acSysCmdInitMeter**; this argument isn't valid for other action argument values.
 **Note**  When you specify the  **acSysCmdGetObjectState** value for the Action parameter, then you must specify the name of the database object.

|

### Return Value

Variant


## Remarks

For example, if you are building a custom wizard that creates a new form, you can use the  **SysCmd** method to display a progress meter indicating the progress of your wizard as it constructs the form.

By calling the  **SysCmd** method with the various progress meter actions, you can display a progress meter in the status bar for an operation that has a known duration or number of steps, and update it to indicate the progress of the operation.

To display a progress meter in the status bar, you must first call the  **SysCmd** method with the **acSysCmdInitMeter** action argument, and the text and value arguments. When the action argument is **acSysCmdInitMeter**, the value argument is the maximum value of the meter, or 100 percent.

To update the meter to show the progress of the operation, call the  **SysCmd** method with the **acSysCmdUpdateMeter** action argument and the value argument. When the action argument is **acSysCmdUpdateMeter**, the **SysCmd** method uses the value argument to calculate the percentage displayed by the meter. For example, if you set the maximum value to 200 and then update the meter with a value of 100, the progress meter will be half-filled.

You can also change the text that's displayed in the status bar by calling the  **SysCmd** method with the **acSysCmdSetStatus** action argument and the text argument. For example, during a sort you might change the text to "Sorting...". When the sort is complete, you would reset the status bar by removing the text. The text argument can contain approximately 80 characters. Because the status bar text is displayed by using a proportional font, the actual number of characters you can display is determined by the total width of all the characters specified by the text argument.

As you increase the width of the status bar text, you decrease the length of the meter. If the text is longer than the status bar and the action argument is  **acSysCmdInitMeter**, the **SysCmd** method ignores the text and doesn't display anything in the status bar. If the text is longer than the status bar and the action argument is **acSysCmdSetStatus**, the **SysCmd** method truncates the text to fit the status bar.

You can't set the status bar text to a zero-length string (" "). If you want to remove the existing text from the status bar, set the text argument to a single space. The following examples illustrate ways to remove the text from the status bar:




```
varReturn = SysCmd(acSysCmdInitMeter, " ", 100) 
varReturn = SysCmd(acSysCmdSetStatus, " ")
```

If the progress meter is already displayed when you set the text by calling the  **SysCmd** method with the **acSysCmdSetStatus** action argument, the **SysCmd** method automatically removes the meter.

Call the  **SysCmd** method with other actions to determine system information about Microsoft Access, including which version number of Microsoft Access is running, whether it is a run-time version, the location of the Microsoft Access executable file, the setting for the /profile argument specified in the command line, and the name of an .ini file associated with Microsoft Access.


 **Note**  Both general and customized settings for Microsoft Access are now stored in the Windows Registry, so you probably won't need an .ini file with your Microsoft Access application. The  **acSysCmdIniFile** action argument exists for compatibility with earlier versions of Microsoft Access.

Call the  **SysCmd** method with the **acSysCmdGetObjectState** action argument and the objecttype and objectname arguments to return the state of a specified database object. An object can be in one of four possible states: not open or nonexistent, open, new, or changed but not saved.

For example, if you are designing a wizard that inserts a new field in a table, you need to determine whether the structure of the table has been changed but not yet saved, so that you can save it before modifying its structure. You can check the value returned by the  **SysCmd** method to determine the state of the table.

The  **SysCmd** method with the **acSysCmdGetObjectState** action argument can return any combination of the following constants.



|**Constant**|**State of database object**|**Value**|
|:-----|:-----|:-----|
|**acObjStateOpen**|Open|1|
|**acObjStateDirty**|Design changed but not saved|2|
|**acObjStateNew**|New|4|

 **Note**  If the object referred to by the objectname argument is either not open or doesn't exist, the  **SysCmd** method returns a value of zero.

The following code can be used to enable the use of your ActiveX control in expressions when the ActiveX control has been added to a form:




```vb
SysCmd 14, "<ActiveX Control GUID>" 
```


 **Note**  Replace <ActiveX Control GUID> with the globally unique identifier (GUID) that identifies the ActiveX control that you want to enable in expressions.


 **Note**  You cannot remove an ActiveX control once it has been added to the list of allowed controls.


## See also


#### Concepts


[Application Object](application-object-access.md)

