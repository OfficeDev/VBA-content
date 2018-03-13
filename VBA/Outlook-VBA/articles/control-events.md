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



Most Microsoft Forms 2.0 controls in an Microsoft Outlook custom form support only one event, the  <strong>Click</strong> event.
A control bound to a field does not fire the  
<strong>Click</strong> event. You must handle the appropriate [field event](field-events.md) to detect a user's interaction with a control bound to a field.
The following Forms 2.0 controls and Outlook controls fire the  <strong>Click</strong> event whenever a user clicks anywhere in the control.<br>
 
<strong><a href="checkbox-object-outlook-forms-script.md" data-raw-source="[CheckBox](checkbox-object-outlook-forms-script.md)">CheckBox</a></strong><br>
 
<strong><a href="commandbutton-object-outlook-forms-script.md" data-raw-source="[CommandButton](commandbutton-object-outlook-forms-script.md)">CommandButton</a></strong><br>
 
<strong><a href="frame-object-outlook-forms-script.md" data-raw-source="[Frame](frame-object-outlook-forms-script.md)">Frame</a></strong><br>
 
<strong><a href="image-object-outlook-forms-script.md" data-raw-source="[Image](image-object-outlook-forms-script.md)">Image</a></strong><br>
 
<strong><a href="label-object-outlook-forms-script.md" data-raw-source="[Label](label-object-outlook-forms-script.md)">Label</a></strong><br>
 
<strong><a href="optionbutton-object-outlook-forms-script.md" data-raw-source="[OptionButton](optionbutton-object-outlook-forms-script.md)">OptionButton</a></strong><br>
 
<strong><a href="togglebutton-object-outlook-forms-script.md" data-raw-source="[ToggleButton](togglebutton-object-outlook-forms-script.md)">ToggleButton</a></strong><br>
 
<strong><a href="olkbusinesscardcontrol-object-outlook.md" data-raw-source="[OlkBusinessCardControl](olkbusinesscardcontrol-object-outlook.md)">OlkBusinessCardControl</a></strong><br>
 
<strong><a href="olkcategory-object-outlook.md" data-raw-source="[OlkCategory](olkcategory-object-outlook.md)">OlkCategory</a></strong><br>
 
<strong><a href="olkcheckbox-object-outlook.md" data-raw-source="[OlkCheckBox](olkcheckbox-object-outlook.md)">OlkCheckBox</a></strong><br>
 
<strong><a href="olkcommandbutton-object-outlook.md" data-raw-source="[OlkCommandButton](olkcommandbutton-object-outlook.md)">OlkCommandButton</a></strong><br>
 
<strong><a href="olkcontactphoto-object-outlook.md" data-raw-source="[OlkContactPhoto](olkcontactphoto-object-outlook.md)">OlkContactPhoto</a></strong><br>
 
<strong><a href="olkdatecontrol-object-outlook.md" data-raw-source="[OlkDateControl](olkdatecontrol-object-outlook.md)">OlkDateControl</a></strong><br>
 
<strong><a href="olkframeheader-object-outlook.md" data-raw-source="[OlkFrameHeader](olkframeheader-object-outlook.md)">OlkFrameHeader</a></strong><br>
 
<strong><a href="olkinfobar-object-outlook.md" data-raw-source="[OlkInfoBar](olkinfobar-object-outlook.md)">OlkInfoBar</a></strong><br>
 
<strong><a href="olklabel-object-outlook.md" data-raw-source="[OlkLabel](olklabel-object-outlook.md)">OlkLabel</a></strong><br>
 
<strong><a href="olkoptionbutton-object-outlook.md" data-raw-source="[OlkOptionButton](olkoptionbutton-object-outlook.md)">OlkOptionButton</a></strong><br>
 
<strong><a href="olksenderphoto-object-outlook.md" data-raw-source="[OlkSenderPhoto](olksenderphoto-object-outlook.md)">OlkSenderPhoto</a></strong><br>
 
<strong><a href="olktextbox-object-outlook.md" data-raw-source="[OlkTextBox](olktextbox-object-outlook.md)">OlkTextBox</a></strong><br>
 
<strong><a href="olktimecontrol-object-outlook.md" data-raw-source="[OlkTimeControl](olktimecontrol-object-outlook.md)">OlkTimeControl</a></strong><br>
 
<strong><a href="olktimezonecontrol-object-outlook.md" data-raw-source="[OlkTimeZoneControl](olktimezonecontrol-object-outlook.md)">OlkTimeZoneControl</a></strong><br>


The following controls fire the  <strong>Click</strong> event when the user selects an item in the list.<br>
 
<strong><a href="combobox-object-outlook-forms-script.md" data-raw-source="[ComboBox](combobox-object-outlook-forms-script.md)">ComboBox</a></strong><br>
 
<strong><a href="listbox-object-outlook-forms-script.md" data-raw-source="[ListBox](listbox-object-outlook-forms-script.md)">ListBox</a></strong><br>
 
<strong><a href="olkcombobox-object-outlook.md" data-raw-source="[OlkComboBox](olkcombobox-object-outlook.md)">OlkComboBox</a></strong><br>
 
<strong><a href="olklistbox-object-outlook.md" data-raw-source="[OlkListBox](olklistbox-object-outlook.md)">OlkListBox</a></strong><br>


The following controls do not support the  <strong>Click</strong> event.<br>
 
<strong><a href="scrollbar-object-outlook-forms-script.md" data-raw-source="[ScrollBar](scrollbar-object-outlook-forms-script.md)">ScrollBar</a></strong><br>
 
<strong><a href="spinbutton-object-outlook-forms-script.md" data-raw-source="[SpinButton](spinbutton-object-outlook-forms-script.md)">SpinButton</a></strong><br>
 
<strong><a href="tabstrip-object-outlook-forms-script.md" data-raw-source="[TabStrip](tabstrip-object-outlook-forms-script.md)">TabStrip</a></strong><br>
 
<strong><a href="textbox-object-outlook-forms-script.md" data-raw-source="[TextBox](textbox-object-outlook-forms-script.md)">TextBox</a></strong><br>


While the  <strong>MultiPage</strong> control itself does not support the <strong>Click</strong> event, an individual <strong><a href="page-object-outlook-forms-script.md" data-raw-source="[Page](page-object-outlook-forms-script.md)">Page</a></strong> on a <strong>MultiPage</strong> control will fire the <strong>Click</strong> event if the user clicks inside the client area of the page, but not if the user clicks the tab associated with the page.<br>


To detect a change in a  **TextBox** control, bind the control to a field and then handle the appropriate field event.
If you have to further extend controls in a custom form, customize a form with Outlook controls in a form region instead of Forms 2.0 controls in a form page. For more information, see  [Controls in a Custom Form](controls-in-a-custom-form.md).

