---
title: Specify Locale-Specific User Interface for a Form Region
ms.prod: outlook
ms.assetid: 74cf8452-5e75-c939-2bf8-91607241bb76
ms.date: 06/08/2017
---


# Specify Locale-Specific User Interface for a Form Region

You can specify localized versions of certain pieces of user interface in a form region. For example, you can specify an English version of the form region display name and a Spanish version of the display name.

To support localized user interface for the same locale, specify the locale and the localized user interface under the  **stringOverride** element in the form region manifest XML file. In general, if there is one or more locales that share the same localization strings, you can specify their locales and localized user interface under the same **stringOverride** element.

Under the  **stringOverride** element, you can specify the localization information in one of two ways: you can either specify it as values of child elements, or specify a localization file that contains the localization information. You will specify the **stringOverride** element in the form region manifest XML file that you will use when you register the form region in the Windows registry. For more information on registering a form region, see [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).

The following table shows the pieces of user interface in a form region that you can localize. If you choose to specify the localization information inline, you can use the corresponding XML child elements under the  **stringOverride** element.


| **Localizeable User Interface**| **XML Child Elements**|
|Display name of the form region| **title**|
|Form region identifier| **formRegionName**|
|Description of the form region| **description**|
|Display name of a control in the form region| **control**|
|Display name and subject prefix of an action defined for the form region| **action**|
If you choose to provide a localization file, you will specify it as the value of the  **file** attribute of the **stringOverride** element.
Localization files follow an XML schema consisting of elements similar to the child elements of the  **stringOverride** element in the form region XML schema. For more information on the localization XML schema, see the Microsoft Outlook 2010 XML Schema Reference in the [MSDN Library](http://msdn.microsoft.com/library).

## To specify the locale


- In the form region manifest XML file, specify the Locale ID (LCID) of the locale as the value of the  **language** attribute of the **stringOverride** element.
    
    If there are multiple locales sharing the same localized user interface, specify the LCID of all these locales, separating them with semicolons, as the combined value of the  **language** attribute of a single **stringOverride** element. Specify only one **stringOverride** element in a form region manifest XML file for each unique value (or combined value) of the **language** attribute.
    
    You can also specify  `all` as the value of the language attribute to specify that the user interface specified in that **stringOverride** element applies to all locales.
    
    The following example lists the locale IDs for Spanish-Spain and French-France as the two locales under the same  **stringOverride** element:
    


```
  <stringOverride language="1034;1036">
    <!-- specify localization strings here -->
</stringOverride>
```


## Optional: To specify a localization file


- In the form region manifest XML file, specify the path to a file that contains the localized strings for the specified locale as the value of the  **file** attribute of the **stringOverride** element.
    
    The path to the localization file can be a full path or a path relative to the location of the form region manifest XML file that you specify when you register the form region.
    
    When specifying the location of the localization file, you can use the system variable  `%langid%` as a placeholder for the LCID of the current **stringOverride** element. For example:
    


```
  <stringOverride language="1034" file="%langid%\UserStrings.xml" />
    <!-- no need to specify localization strings here -->
```


     **Note**  If a file is specified for the  **file** attribute of the **stringOverride** element, Outlook will ignore all child elements of the **stringOverride** element.

## Optional: To specify a localization string for the display name of the form region


- In the form region manifest XML file, under the  **stringOverride** element, specify a string identifier for the form region as a value of the **title** child element.
    
    The value of the  **title** element is the display name of the form region for the specified locale or locales. If the form region is a replacement or replace-all form region, the value of **title** is displayed in the **Actions** menu and the **Choose Form** dialog box.
    
    The following example specifies a display name for a form region localized for the English-Canada locale.
    


```
  <stringOverride language="4105">
    <title>Template for Canadians</title>
</stringOverride>
```


## Optional: To specify a localization string for the form region identifier


- In the form region manifest XML file, under the  **stringOverride** element, specify a string identifier for the form region as a value of the **formRegionName** child element.
    
    The value of the  **formRegionName** element identifies the form region in the **Show** tab of the ribbon for the specified locale or locales. If the form region is an adjoining form region, the value is also used in the header that separates the beginning of an adjoining form region from the preceding portion of the form.
    
    The following example specifies  `Addendum` as the form region identifier of an adjoining form region localized for the English-Canada locale:
    


```
  <stringOverride language="4105">
    <formRegionName>Addendum</formRegionName>
</stringOverride>

```


## Optional: To specify a localization string for the description of the form region


- In the form region manifest XML file, under the  **stringOverride** element, specify a string identifier for the form region as a value of the **description** child element.
    
    The value of the  **description** element is a text description of the form region for the specified locale or locales. If the form region is a replacement or replace-all form region, the value of **description** is displayed in the **Choose Form** dialog box.
    
    The following example specifies a description for a form region localized for the English-Canada locale:
    


```
  <stringOverride language="4105">
    <description>This template is intended for English speaking Canadians.</description>
</stringOverride>

```


## Optional: To specify a localization string for the display name of a control in the form region


1. In the form region manifest XML file, under the  **stringOverride** element, specify the value of the name property as the value of the **name** attribute of the **control** child element.
    
    The value of the name property is accessible from the user interface when you right-click the control in the Forms Designer and select  **Properties**. For example, the default name for an Outlook Text Box control is  **TextBox1**.
    
2. Under the  **control** element, specify a localized string for the **caption** child element.
    
    The value of the  **caption** element is the display name of the control localized for the specified locale or locales.
    
    The following example specifies a localized name for an Outlook Text Box control for the English-Canada locale:
    


```
  <stringOverride language="4105">
    <control name="TextBox1">
        <caption>Canadian postal code</caption>
    </control>
</stringOverride>
```


## Optional: To specify a localization string for the display name of an action


1. In the form region manifest XML file, under the  **stringOverride** element, specify the internal name of the action as the value of the **name** attribute of the **action** child element.
    
2. Under the  **action** element, specify a localized string for the display name of the action as the value of the **title** child element.
    
    The value of the  **title** element is localized for the specified locale or locales.
    
    The following example specifies a localized display name for a custom action for the English-Canada locale:
    


```
  <stringOverride language="4105">
    <action name="replyToBlog">
        <title>Reply to Blog</title>
    </action>
</stringOverride>

```


## Optional: To specify a localization string for the subject prefix of an item resulting from an action


1. In the form region manifest XML file, under the  **stringOverride** element, specify the internal name of the action as the value of the **name** attribute of the **action** child element.
    
2. Under the  **action** element, specify a localized string for the subject prefix as the value of the **subject** child element.
    
    The value of the  **subject** element is the prefix for the subject field of an item that results from the action, and is localized for the specified locale or locales.
    
    The following example specifies a localized subject prefix for a custom action for the English-Canada locale:
    


```
  <stringOverride language="4105">
    <action name="replyToBlog">
        <subject>Re</subject>
    </action>
</stringOverride>

```


