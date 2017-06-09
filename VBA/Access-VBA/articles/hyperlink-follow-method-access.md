---
title: Hyperlink.Follow Method (Access)
keywords: vbaac10.chm10117
f1_keywords:
- vbaac10.chm10117
ms.prod: access
api_name:
- Access.Hyperlink.Follow
ms.assetid: 842f546c-b629-fd47-e8d0-d73d3ee7f3cd
ms.date: 06/08/2017
---


# Hyperlink.Follow Method (Access)

The  **Follow** method opens the document or Web page specified by a hyperlink address associated with a control on a form or report.


## Syntax

 _expression_. **Follow**( ** _NewWindow_**, ** _AddHistory_**, ** _ExtraInfo_**, ** _Method_**, ** _HeaderInfo_** )

 _expression_ A variable that represents a **Hyperlink** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewWindow_|Optional|**Boolean**|**True** (?1) opens the document in a new window and **False** (0) opens the document in the current window. The default is **False**.|
| _AddHistory_|Optional|**Boolean**|**True** adds the hyperlink to the History folder and **False** doesn't add the hyperlink to the History folder. The default is **True**.|
| _ExtraInfo_|Optional|**Variant**|A string or an array of Byte data that specifies additional information for navigating to a hyperlink. For example, this argument may be used to specify a search parameter for an .ASP or .IDC file. In your Web browser, the  _extrainfo_ argument may appear after the hyperlink address, separated from the address by a question mark (?). You don't need to include the question mark when you specify the _extrainfo_ argument.|
| _Method_|Optional|**MsoExtraInfoMethod**|A  **[MsoExtraInfoMethod](http://msdn.microsoft.com/library/eb8edb9c-2a9a-62b5-f592-e40a2325a555%28Office.15%29.aspx)** constant that specifies how the _extrainfo_ argument is attached. The default is **msoMethodGet**.|
| _HeaderInfo_|Optional|**String**|Specifies header information. By default the  _headerinfo_ argument is a zero-length string (" ").|

## Remarks

The  **Follow** method has the same effect as clicking a hyperlink.

You can include the  **Follow** method in an event procedure if you want to open a hyperlink in response to a user action. For example, you may want to open a web page with reference information when a user opens a particular form.

When you use the  **Follow** method, you don't need to know the address specified by a control's **HyperlinkAddress** property. You only need to know the name of the control that contains the hyperlink. Conversely, when you use the **[FollowHyperlink](application-followhyperlink-method-access.md)** method, you need to specify the address for the particular hyperlink you wish to follow.


## Example

The following example sets the  **HyperlinkAddress** property of a command button and then opens the hyperlink when the form is loaded.

To try this example, create a form and add a command button named Command0. Paste the following code into the form's module and switch to Form view:




```vb
Private Sub Form_Load() 
    Dim ctl As CommandButton 
 
    Set ctl = Me!Command0 
    With ctl 
        .Visible = False 
        .HyperlinkAddress = "http://www.microsoft.com/" 
        .Hyperlink.Follow 
    End With 
End Sub
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-access.md)

