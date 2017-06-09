---
title: Specify the Location of a Form Region in a Custom Form
ms.prod: outlook
ms.assetid: c617f6a3-c39a-bb0f-37ff-1ea999dac8be
ms.date: 06/08/2017
---


# Specify the Location of a Form Region in a Custom Form

A form region is a piece of custom user interface that you add to a form. You can designate the form region to be displayed in one of several ways in a custom form. To do so, you will be specifying the  **formRegionType** and **displayAfter** elements of the form region XML schema in the corresponding form region manifest XML file.


## On the Default Page

There are several ways you can display a form region or form regions on the default page of a standard form.


### To add a form region to the default page




- In the form region manifest XML file, specify  **adjoining** as the value of the **formRegionType** element.
    
The form region will be added to the bottom of the default page of the original standard form, and will be displayed in an Inspector or the Reading Pane. 

For example, to add a form region to the bottom of the default page of the standard Message form, you can specify the following in the form region manifest XML file of the form region:




```
<formRegionType>adjoining</formRegionType>
```

You can use the resulting custom form to display items of the same message class as the original standard form, or you can assign a derived message class for the custom form and use the custom form to display only items of the derived message class.


### To add multiple form regions to the default page


1. For each form region, in the corresponding form region manifest XML file, specify  **adjoining** as the value of the **formRegionType** element.
    
2. Except for the form region that will appear as the first form region on the default page, for each of the other form regions, in the corresponding form region manifest XML file, specify the internal name of the form region that will precede this one as the value of the  **displayAfter** element.
    
You can use the resulting custom form to display items of the same message class as the original standard form, or you can assign a derived message class for the custom form and use the custom form to display only items of the derived message class.

The first form region will be added to the bottom of the default page of the original standard form, and will be appended by the other form regions in the order that you have specified in the corresponding  **displayAfter** element.

For example, if you want to order three form regions, A, B, and C, that have the internal names  **FormRegionA**,  **FormRegionB**, and  **FormRegionC** to be displayed in the order A, B, and C, you will specify the following in A's form region manifest XML file:




```
<formRegionType>adjoining</formRegionType>
```

You will specify the following in B's form region manifest XML file:




```
<formRegionType>adjoining</formRegionType> 
<displayAfter>FormRegionA</displayAfter>
```

You will specify the following in C's form region manifest XML file:




```
<formRegionType>adjoining</formRegionType> 
<displayAfter>FormRegionB</displayAfter>
```


 **Note**  You can use the  **displayAfter** element to specify the order of multiple adjoining form regions in a custom form. However, this order is only valid the first time that the form is displayed for the user on the local computer. The user has the option to change the order of adjoining form regions by opening the form and moving the form regions up or down on the default page through the form region header context menu. Outlook caches the updated order and uses the cached order on subsequent displays of the form.


### To "replace" the entire default page by a form region


1. In the form region manifest XML file, specify  **replace** as the value of the **formRegionType** element.
    
2. When you register the form region in the Windows registry, under the local machine key (as  **HKEY_LOCAL_MACHINE\Software\Microsoft\Office\Outlook\FormRegions**) or the current user key (as  **HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\FormRegions**), create a separate key for the derived message class for this form region. Add a value of the type  **REG_SZ**, specifying the internal name of the form region as the name of the key, and the full local file path name to the form region manifest XML file as the data of the key.
    
When you are "replacing" the default page of a standard form, you are in reality using the standard form as a template and creating a new custom form that has your form region as the default page. If the original standard form contains other form pages or separate form regions, they will remain as part of the custom form.

You must assign a derived message class to the resulting custom form and use the form to display items of that message class.

For example, you have created a form region that has the internal name CustomPage and form region manifest XML file CustomPage.xml in c:\Form Regions\. To use the form region to replace the default page of the standard Message form, you can specify the following in the form region manifest XML file of CustomPage:




```
<formRegionType>replace</formRegionType>
```

When you register this form region in the Windows registry, you must not specify the message class of the original standard form,  **IPM.Note**, but specify a derived message class, such as  **IPM.Note.CustomPage**. For this example, you will register the form region under the current user key,  **HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\FormRegions**, by creating a key  **IPM.Note.CustomPage**. You will then add a value of the type  **REG_SZ**, specifying the internal name,  **CustomPage**, of the form region as the name of the key, and the full local file path name to the form region manifest XML file,  **c:\Form Regions\CustomPage.xml**, as the data of the key.


### To "replace" the entire standard form by a form region


1. In the form region manifest XML file, specify  **replaceall** as the value of the **formRegionType** element.
    
2. When you register the form region in the Windows registry, under the local machine key (as  **HKEY_LOCAL_MACHINE\Software\Microsoft\Office\Outlook\FormRegions**) or the current user key (as  **HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\FormRegions**), create a separate key for the derived message class for this form region. Add a value of the type  **REG_SZ**, specifying the internal name of the form region as the name of the key, and the full local file path name to the form region manifest XML file as the data of the key.
    
When you are "replacing" the entire standard form with a form region, you are in reality using the standard form as a template and creating a new custom form that has the form region as the default page. If the original standard form contains other form pages or separate form regions, they will not remain as part of the custom form.

You must assign a derived message class to the resulting custom form and use the form to display items of that message class.

For example, you have created a form region that has the internal name CustomMessage and form region manifest XML file CustomMessage.xml in c:\Form Regions\. To use the standard Message form as the template for a new custom form that will contain CustomMessage as the default page, you can specify the following in the form region manifest XML file of CustomMessage:




```
<formRegionType>replaceall</formRegionType>
```

When you register this form region in the Windows registry, you will specify a derived message class, such as  **IPM.Note.CustomMessage**. For this example, you will register the form region under the current user key,  **HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\FormRegions**, by creating a key  **IPM.Note.CustomMessage**. You will then add a value of the type  **REG_SZ**, specifying the internal name,  **CustomMessage**, of the form region as the name of the key, and the full local file path name to the form region manifest XML file,  **c:\Form Regions\CustomMessage.xml**, as the data of the key.


## On Pages Other than the Default Page

You can add one or more form regions as separate pages to a standard form.


### To add a form region as a separate page


- In the form region manifest XML file, specify  **separate** as the value of the **formRegionType** element.
    
The form region will be added as a separate page following all the existing pages of the original standard form, and will be displayed as a standalone page in an Inspector.

For example, to add a form region as a separate page to the standard Contact form, you can specify the following in the form region manifest XML file of the form region:




```
<formRegionType>separate</formRegionType>
```

The form region will be displayed as a separate page following the  **All Fields** page of the standard Contact form.

You can use the resulting custom form to display items of the same message class as the original standard form, or you can assign a derived message class for the custom form and use the custom form to display only items of the derived message class.


### To add multiple form regions as separate pages


1. For each form region, in the corresponding form region manifest XML file, specify  **separate** as the value of the **formRegionType** element.
    
2. Except for the form region that will appear as the first form region in the custom form, for each of the other form regions, in the corresponding form region manifest XML file, specify the internal name of the form region that will precede this one as the value of the  **displayAfter** element.
    
You can use the resulting custom form to display items of the same message class as the original standard form, or you can assign a derived message class for the custom form and use the custom form to display only items of the derived message class.

The first form region will be added as a separate page of the original standard form, and will be appended by the other form regions in the order that you have specified in the corresponding  **displayAfter** element.

For example, if you want to order three separate form regions, A, B, and C, that have the internal names  **FormRegionA**,  **FormRegionB**, and  **FormRegionC** to be displayed as separate pages in the order A, B, and C, you will specify the following in A's form region manifest XML file:




```
<formRegionType>separate</formRegionType>
```

You will specify the following in B's form region manifest XML file:




```
<formRegionType>separate</formRegionType>
<displayAfter>FormRegionA</displayAfter>
```

You will specify the following in C's form region manifest XML file:




```
<formRegionType>separate</formRegionType>
<displayAfter>FormRegionB</displayAfter>
```


