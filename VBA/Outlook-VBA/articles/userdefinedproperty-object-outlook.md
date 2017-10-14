---
title: UserDefinedProperty Object (Outlook)
keywords: vbaol11.chm3151
f1_keywords:
- vbaol11.chm3151
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperty
ms.assetid: aebe38db-0ff9-79d2-b5a7-751fea7c97f3
ms.date: 06/08/2017
---


# UserDefinedProperty Object (Outlook)

Represents the definition of a user-defined property for a  **[Folder](folder-object-outlook.md)** object.


## Remarks

Use  **[UserDefinedProperties](folder-userdefinedproperties-property-outlook.md)** ( _index_), where  _index_ is a name or index number, to return a single **UserDefinedProperty** object.

Use the  **[Add](userdefinedproperties-add-method-outlook.md)** method of the **[UserDefinedProperties](folder-userdefinedproperties-property-outlook.md)** collection for a **Folder** object to define a user-defined property for that folder.

Use the  **[Type](userdefinedproperty-type-property-outlook.md)** property to return the user-defined property type and the **[DisplayFormat](userdefinedproperty-displayformat-property-outlook.md)** property to return the display format for the user-defined property. If the **Type** property is set to **olCombination** or **olFormula**, use the **[Formula](userdefinedproperty-formula-property-outlook.md)** property to return the formula used to generate values for the user-defined property.

The  **UserDefinedProperty** object represents only the definition of a user-defined property, which is applicable to all Outlook items contained by the folder. To retrieve or change user-defined property values for an Outlook item in that folder, use the **[UserProperties](mailitem-userproperties-property-outlook.md)** property of the Outlook item, such as a **[MailItem](mailitem-object-outlook.md)** object, to retrieve the **[UserProperties](userproperties-object-outlook.md)** collection for that item. You can then use the **[UserProperty](userproperty-object-outlook.md)** object for the appropriate user-defined property to retrieve or change the value of that user-defined property for the Outlook item.


## Example

The following Visual Basic for Applications (VBA) example displays the name of a specified  **Folder** object, as well as the name and type of every **UserDefinedProperty** object contained in the **UserDefinedProperties** collection of the specified **Folder** object, to the **Immediate** window.


```
Sub DisplayUserProperties(ByRef FolderToCheck As Folder) 
 Dim objProperty As UserDefinedProperty 
 
 ' Print the name of the specified Folder object 
 ' reference to the Immediate window. 
 Debug.Print "--- Folder: " &amp; FolderToCheck.Name 
 
 ' Check if there are any user-defined properties 
 ' associated with the Folder object reference. 
 If FolderToCheck.UserDefinedProperties.Count = 0 Then 
 ' No user-defined properties are present. 
 Debug.Print " No user-defined properties." 
 Else 
 ' Iterate through every user-defined property in 
 ' the folder. 
 For Each objProperty In FolderToCheck.UserDefinedProperties 
 ' Retrieve the name of the user-defined property. 
 strPropertyInfo = objProperty.Name 
 ' Retrieve the type of the user-defined property. 
 Select Case objProperty.Type 
 Case OlUserPropertyType.olCombination 
 strPropertyInfo = strPropertyInfo &amp; " (Combination)" 
 Case OlUserPropertyType.olCurrency 
 strPropertyInfo = strPropertyInfo &amp; " (Currency)" 
 Case OlUserPropertyType.olDateTime 
 strPropertyInfo = strPropertyInfo &amp; " (Date/Time)" 
 Case OlUserPropertyType.olDuration 
 strPropertyInfo = strPropertyInfo &amp; " (Duration)" 
 Case OlUserPropertyType.olEnumeration 
 strPropertyInfo = strPropertyInfo &amp; " (Enumeration)" 
 Case OlUserPropertyType.olFormula 
 strPropertyInfo = strPropertyInfo &amp; " (Formula)" 
 Case OlUserPropertyType.olInteger 
 strPropertyInfo = strPropertyInfo &amp; " (Integer)" 
 Case OlUserPropertyType.olKeywords 
 strPropertyInfo = strPropertyInfo &amp; " (Keywords)" 
 Case OlUserPropertyType.olNumber 
 strPropertyInfo = strPropertyInfo &amp; " (Number)" 
 Case OlUserPropertyType.olOutlookInternal 
 strPropertyInfo = strPropertyInfo &amp; " (Outlook Internal)" 
 Case OlUserPropertyType.olPercent 
 strPropertyInfo = strPropertyInfo &amp; " (Percent)" 
 Case OlUserPropertyType.olSmartFrom 
 strPropertyInfo = strPropertyInfo &amp; " (Smart From)" 
 Case OlUserPropertyType.olText 
 strPropertyInfo = strPropertyInfo &amp; " (Text)" 
 Case OlUserPropertyType.olYesNo 
 strPropertyInfo = strPropertyInfo &amp; " (Yes/No)" 
 Case Else 
 strPropertyInfo = strPropertyInfo &amp; " (Unknown)" 
 End Select 
 
 ' Print the name and type of the user-defined property 
 ' to the Immediate window. 
 Debug.Print strPropertyInfo 
 Next 
 End If 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[Delete](userdefinedproperty-delete-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](userdefinedproperty-application-property-outlook.md)|
|[Class](userdefinedproperty-class-property-outlook.md)|
|[DisplayFormat](userdefinedproperty-displayformat-property-outlook.md)|
|[Formula](userdefinedproperty-formula-property-outlook.md)|
|[Name](userdefinedproperty-name-property-outlook.md)|
|[Parent](userdefinedproperty-parent-property-outlook.md)|
|[Session](userdefinedproperty-session-property-outlook.md)|
|[Type](userdefinedproperty-type-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
