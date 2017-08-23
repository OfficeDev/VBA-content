---
title: "Свойство Hyperlink.Address (издатель)"
keywords: vbapb10.chm4587523
f1_keywords: vbapb10.chm4587523
ms.prod: publisher
api_name: Publisher.Hyperlink.Address
ms.assetid: 784a9213-38bc-c5fd-f215-abeb174ec628
ms.date: 06/08/2017
ms.openlocfilehash: c61f9f85d591a563ba567ee206ad8ae5c2357b0e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinkaddress-property-publisher"></a>Свойство Hyperlink.Address (издатель)

Возвращает или задает **строку** , представляющую URL-адрес гиперссылки. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Адрес**

 переменная _expression_A, представляющий объект **гиперссылки** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере отображается URL-адреса для всех гиперссылок в активной публикации.


```vb
Sub ShowHyperlinkAddresses() 
 Dim pgsPage As Page 
 Dim shpShape As Shape 
 Dim hprLink As Hyperlink 
 Dim intCount As Integer 
 For Each pgsPage In ActiveDocument.Pages 
 For Each shpShape In pgsPage.Shapes 
 If shpShape.TextFrame.TextRange.Hyperlinks.Count > 0 Then 
 For Each hprLink In shpShape.TextFrame.TextRange.Hyperlinks 
 MsgBox "This hyperlink goes to " &; hprLink.Address &; "." 
 intCount = intCount + 1 
 Next hprLink 
 ElseIf shpShape.Hyperlink.Address <> "" Then 
 MsgBox "This hyperlink goes to " &; shpShape.Hyperlink.Address &; "." 
 intCount = intCount + 1 
 End If 
 Next shpShape 
 Next pgsPage 
 If intCount < 1 Then 
 MsgBox "You don't have any hyperlinks in your publication." 
 Else 
 MsgBox "You have " &; intCount &; " hyperlinks in " &; ThisDocument.Name &; "." 
 End If 
End Sub
```


