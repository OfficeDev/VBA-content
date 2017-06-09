---
title: Control Events
keywords: olfm10.chm3077124
f1_keywords:
- olfm10.chm3077124
ms.prod: outlook
ms.assetid: 6305af2d-d26c-024f-945a-8eaa773bab85
ms.date: 06/08/2017
---


# Control Events



Most Microsoft Forms 2.0 controls in an Microsoft Outlook custom form support only one event, the  **Click** event.
A control bound to a field does not fire the  **Click** event. You must handle the appropriate [field event](field-events.md) to detect a user's interaction with a control bound to a field.
The following Forms 2.0 controls and Outlook controls fire the  **Click** event whenever a user clicks anywhere in the control.<br>
 **[CheckBox](checkbox-object-outlook-forms-script.md)**<br>
 **[CommandButton](commandbutton-object-outlook-forms-script.md)**<br>
 **[Frame](frame-object-outlook-forms-script.md)**<br>
 **[Image](image-object-outlook-forms-script.md)**<br>
 **[Label](label-object-outlook-forms-script.md)**<br>
 **[OptionButton](optionbutton-object-outlook-forms-script.md)**<br>
 **[ToggleButton](togglebutton-object-outlook-forms-script.md)**<br>
 **[OlkBusinessCardControl](olkbusinesscardcontrol-object-outlook.md)**<br>
 **[OlkCategory](olkcategory-object-outlook.md)**<br>
 **[OlkCheckBox](olkcheckbox-object-outlook.md)**<br>
 **[OlkCommandButton](olkcommandbutton-object-outlook.md)**<br>
 **[OlkContactPhoto](olkcontactphoto-object-outlook.md)**<br>
 **[OlkDateControl](olkdatecontrol-object-outlook.md)**<br>
 **[OlkFrameHeader](olkframeheader-object-outlook.md)**<br>
 **[OlkInfoBar](olkinfobar-object-outlook.md)**<br>
 **[OlkLabel](olklabel-object-outlook.md)**<br>
 **[OlkOptionButton](olkoptionbutton-object-outlook.md)**<br>
 **[OlkSenderPhoto](olksenderphoto-object-outlook.md)**<br>
 **[OlkTextBox](olktextbox-object-outlook.md)**<br>
 **[OlkTimeControl](olktimecontrol-object-outlook.md)**<br>
 **[OlkTimeZoneControl](olktimezonecontrol-object-outlook.md)**<br>
 
The following controls fire the  **Click** event when the user selects an item in the list.<br>
 **[ComboBox](combobox-object-outlook-forms-script.md)**<br>
 **[ListBox](listbox-object-outlook-forms-script.md)**<br>
 **[OlkComboBox](olkcombobox-object-outlook.md)**<br>
 **[OlkListBox](olklistbox-object-outlook.md)**<br>

The following controls do not support the  **Click** event.<br>
 **[ScrollBar](scrollbar-object-outlook-forms-script.md)**<br>
 **[SpinButton](spinbutton-object-outlook-forms-script.md)**<br>
 **[TabStrip](tabstrip-object-outlook-forms-script.md)**<br>
 **[TextBox](textbox-object-outlook-forms-script.md)**<br>

While the  **MultiPage** control itself does not support the **Click** event, an individual **[Page](page-object-outlook-forms-script.md)** on a **MultiPage** control will fire the **Click** event if the user clicks inside the client area of the page, but not if the user clicks the tab associated with the page.<br>

To detect a change in a  **TextBox** control, bind the control to a field and then handle the appropriate field event.
If you have to further extend controls in a custom form, customize a form with Outlook controls in a form region instead of Forms 2.0 controls in a form page. For more information, see  [Controls in a Custom Form](controls-in-a-custom-form.md).

