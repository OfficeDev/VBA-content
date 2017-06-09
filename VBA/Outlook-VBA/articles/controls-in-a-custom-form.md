---
title: Controls in a Custom Form
ms.prod: outlook
ms.assetid: fcba1b34-c526-5d01-8644-cb8852bd2348
ms.date: 06/08/2017
---


# Controls in a Custom Form

Form controls allow you to customize the user interface and behavior of a custom form. Because all code behind a form is called from a form event or a control event, programming the response to control events is one of the primary means to extend a form programmatically. This topic describes adding and displaying Microsoft Forms 2.0 controls and Microsoft Outlook controls in Outlook forms. For information on programming Forms 2.0 controls, see the Form Script Reference. For more information on programming Outlook controls, see the Object Model Reference.


## Design Time

In the forms designer, you can add a control from the control toolbox to a form page or a form region. The control toolbox is initialized with a number of Forms 2.0 controls. You can add other ActiveX controls that have been installed on your computer to the control toolbox.


## Microsoft Forms 2.0 Controls

The following Forms 2.0 controls exist in the control toolbox by default:


- Microsoft Forms 2.0 CheckBox control
    
-  Microsoft Forms 2.0 ComboBox control
    
- Microsoft Forms 2.0 CommandButton control
    
- Microsoft Forms 2.0 Frame control
    
- Microsoft Forms 2.0 Image control
    
- Microsoft Forms 2.0 Label control
    
- Microsoft Forms 2.0 ListBox control
    
- Microsoft Forms 2.0 MultiPage control
    
- Microsoft Forms 2.0 OptionButton control
    
- Microsoft Forms 2.0 ScrollBar control
    
- Microsoft Forms 2.0 SpinButton control
    
- Microsoft Forms 2.0 TabStrip control
    
- Microsoft Forms 2.0 TextBox control
    
- Microsoft Forms 2.0 ToggleButton control
    

## Microsoft Outlook Controls

The following table lists Outlook controls that are installed on your computer when you install Microsoft Office Outlook 2007 or a later version of Outlook. These controls support Microsoft Windows themes. Before you can use them in Outlook forms, you must first add them to the control toolbox. 

Use these controls only in form regions and not form pages in custom forms.

Some of these controls are designed to leverage features in Outlook, for example, the Microsoft Outlook Body Control is designed to display the body of an Outlook item. All of the Outlook controls are designed for use only in Outlook add-ins. 



| **Control**| **Designed for Specific Outlook Features**|
|:-----|:-----|
|Microsoft Outlook Body Control|Yes|
|Microsoft Outlook Business Card Control|Yes|
|Microsoft Outlook Category Control|Yes|
|Microsoft Outlook Check Box Control|No|
|Microsoft Outlook Combo Box Control|No|
|Microsoft Outlook Command Button Control|No|
|Microsoft Outlook Contact Photo Control|Yes|
|Microsoft Outlook Date Control|Yes|
|Microsoft Outlook Frame Header Control|Yes|
|Microsoft Outlook InfoBar Control|Yes|
|Microsoft Outlook Label Control|No|
|Microsoft Outlook List Box Control|No|
|Microsoft Outlook Option Button Control|No|
|Microsoft Outlook Page Control|Yes|
|Microsoft Outlook Recipient Control|Yes|
|Microsoft Outlook Sender Photo Control|Yes|
|Microsoft Outlook Text Box Control|No|
|Microsoft Outlook Time Control|Yes|
|Microsoft Outlook Time Zone Control|Yes|

## Run Time

On custom form pages, Forms 2.0 controls are always displayed with a classic look without Windows theming. In a form region, any Forms 2.0 control that has a themed Outlook counterpart control assumes an appearance that is themed to Windows and in fact can be cast with the type of the counterpart. For example, if the user has dropped a Forms 2.0 TextBox control in a form region, programmatically, Outlook will replace this instance of the control by an instance of the Outlook counterpart control, and you will be able to apply a cast of  **Microsoft.Office.Interop.Outlook.OlkTextBox** to this control and access it as an Outlook TextBox control. You should not access it as a **Microsoft.Vbe.Interop.Forms.TextBox** control. For more information on casting controls in an add-in, see [Extending a Form Region with an Add-in](extending-a-form-region-with-an-add-in.md). The following table lists each Forms 2.0 control and the corresponding Outlook control (if one exists).



| **Forms 2.0 Control**| **Outlook Control**|
|:-----|:-----|
|Microsoft Forms 2.0 CheckBox control|Microsoft Outlook Check Box Control|
|Microsoft Forms 2.0 ComboBox control|Microsoft Outlook Combo Box Control|
|Microsoft Forms 2.0 CommandButton control|Microsoft Outlook Command Button Control|
|Microsoft Forms 2.0 Frame control|Microsoft Outlook Frame Header Control|
|Microsoft Forms 2.0 Image control| _(No parity)_|
|Microsoft Forms 2.0 Label control|Microsoft Outlook Label Control|
|Microsoft Forms 2.0 ListBox control|Microsoft Outlook List Box Control|
|Microsoft Forms 2.0 MultiPage control| _(No parity)_|
|Microsoft Forms 2.0 OptionButton control|Microsoft Outlook Option Button Control|
|Microsoft Forms 2.0 ScrollBar control| _(No parity)_|
|Microsoft Forms 2.0 SpinButton control| _(No parity)_|
|Microsoft Forms 2.0 TabStrip control| _(No parity)_|
|Microsoft Forms 2.0 TextBox control|Microsoft Outlook Text Box Control|
|Microsoft Forms 2.0 ToggleButton control| _(No parity)_|
Because Outlook controls are installed on computers running Office Outlook 2007 or later, a form that contains these controls will not be displayed properly in any earlier version of Outlook.


## Summary

The following summarizes the differences between Forms 2.0 controls and Outlook controls:



| **Comparison Aspect**| **Forms 2.0 Controls**| **Outlook Controls**|
|:-----|:-----|:-----|
|Available in Outlook 2003 or earlier|Yes|No|
|Available in Office Outlook 2007 or later|Yes|Yes|
|Exists in control toolbox by default|Yes|No|
|How displayed on a form page in Office Outlook 2007 or later|Classic look without Windows theming|Do not use Outlook controls in custom form pages, but use only in form regions|
|How displayed in a form region in Office Outlook 2007 or later|Displayed as its themed counterpart, if one exists, and can be cast with the type of its themed counterpart; classic look if themed counterpart does not exist|Themed look|
|Controls displayed properly in runtime in Outlook 2003 or earlier|Yes|No|
|Controls displayed properly in runtime in Office Outlook 2007 or later|Yes|Yes|



