---
title: "Свойство WebNavigationBarSet.Links (издатель)"
keywords: vbapb10.chm8519697
f1_keywords: vbapb10.chm8519697
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.Links
ms.assetid: 9f155781-390b-ad77-8db7-5099be1409ce
ms.date: 06/08/2017
ms.openlocfilehash: 16c81fccb9427c91bbdced923554639d4772e26a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetlinks-property-publisher"></a>Свойство WebNavigationBarSet.Links (издатель)

Возвращает коллекцию **WebNavigationBarHyperlinks** , содержащую всех гиперссылок в указанный набор панель навигации Web. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Ссылки**

 переменная _expression_A, представляет собой объект- **WebNavigationBarSet** .


### <a name="return-value"></a>Возвращаемое значение

WebNavigationBarHyperlinks


## <a name="example"></a>Пример

Свойство **ссылки** возвращает свойство **WebNavigationBarHyperlinks** . В этом примере возвращается Web гиперссылки первого набора панель навигации Web активного документа.


```vb
ActiveDocument.WebNavigationBarSets(1).Links
```

Пример добавления новой панели навигации веб задайте в активный документ, добавляет гиперссылку на панель навигации и затем добавляет на панели навигации на все страницы публикации, который имеет свойство **AddHyperlinkToWebNavbar** задано значение **True** или свойство **Page.WebPageOptions.IncludePageOnNewWebNavigationBars** задано значение **True**.




```vb
With ActiveDocument.WebNavigationBarSets.AddSet(Name:="WebNavigationBarSet1") 
 With .Links 
 .Add Address:="www.microsoft.com", TextToDisplay:="Microsoft", Index:=1 
 End With 
 .AddToEveryPage Left:=10, Top:=10 
End With
```


