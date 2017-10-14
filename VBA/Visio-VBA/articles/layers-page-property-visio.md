---
title: Layers.Page Property (Visio)
keywords: vis_sdr.chm11913980
f1_keywords:
- vis_sdr.chm11913980
ms.prod: visio
api_name:
- Visio.Layers.Page
ms.assetid: f9fbcbb7-513f-0dc1-a63a-c9936638af4c
ms.date: 06/08/2017
---


# Layers.Page Property (Visio)

Gets the page that contains the layers. Read-only.


## Syntax

 _expression_ . **Page**

 _expression_ A variable that represents a **Layers** object.


### Return Value

Page


## Remarks

If the **Layers** collection is in a master rather than in a page, the **Page** property returns **Nothing** . You cannot set the **Page** property of a **Layers** collection.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Page** property to return a **Page** object from various other objects.


```vb
Public Sub Page_Example() 
 
 Dim vsoPage1 As Visio.Page 
 Dim vsoPage2 As Visio.Page 
 Dim vsoTempPage As Visio.Page 
 Dim vsoLayer1 As Visio.Layer 
 Dim vsoLayer2 As Visio.Layer 
 Dim vsoLayers1 As Visio.Layers 
 Dim vsoLayers2 As Visio.Layers 
 
 'Set the current page name to MyPage1. 
 ActivePage.Name = "MyPage1" 
 
 'Use the Page property to return the current 
 'Page object from the Window object. 
 Set vsoPage1 = ActiveWindow.Page 
 
 'Verify that the expected page was received. 
 Debug.Print "The active window contains: " &; vsoPage1.Name 
 
 'Add a second page named MyPage2. 
 Set vsoPage2 = ActiveDocument.Pages.Add 
 vsoPage2.Name = "MyPage2" 
 
 'Get the Layers collection from each page. 
 Set vsoLayers1 = vsoPage1.Layers 
 Set vsoLayers2 = vsoPage2.Layers 
 
 'Create a layer for each of the Layers collections. 
 Set vsoLayer1 = vsoLayers1.Add("ExampleLayer1") 
 Set vsoLayer2 = vsoLayers2.Add("ExampleLayer2") 
 
 'Use the Page property to return the Page object 
 'from a Layers object. 
 Set vsoTempPage = vsoLayers1.Page 
 
 'Verify that the expected page was received. 
 Debug.Print " vsoLayers1 is from: " &; vsoTempPage.Name 
 
 'Use the Page property to return the Page object 
 'from a Layer object. 
 Set vsoTempPage = vsoLayer2.Page 
 
 'Verify that the expected page was received. 
 Debug.Print " vsoLayer2 is from: " &; vsoTempPage.Name 
 
 'Set the active window's page to "MyPage1." 
 ActiveWindow.Page = "MyPage1" 
 
End Sub
```


