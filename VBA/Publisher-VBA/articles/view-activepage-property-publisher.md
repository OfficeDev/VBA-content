---
title: "Свойство View.ActivePage (издатель)"
keywords: vbapb10.chm327683
f1_keywords: vbapb10.chm327683
ms.prod: publisher
api_name: Publisher.View.ActivePage
ms.assetid: 29289fb2-6692-4cb5-a9e2-b2edb9e9cd7e
ms.date: 06/08/2017
ms.openlocfilehash: a4850678eb5ba7f296ad9705cdf9111bcf7d1084
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="viewactivepage-property-publisher"></a>Свойство View.ActivePage (издатель)

Возвращает объект **[Page](page-object-publisher.md)** , представляющего страницу, отображаемых в окне Microsoft Publisher.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ActivePage**

 переменная _expression_A, представляющий объект **View** .


### <a name="return-value"></a>Возвращаемое значение

Page


## <a name="example"></a>Пример

В этом примере сохраняет активную страницу как изображения JPEG. (Обратите внимание на то, что действительный путь к файлу для работы этого примера необходимо заменить PathToFile.)


```vb
Sub SavePageAsPicture() 
 ActiveView.ActivePage.SaveAsPicture _ 
 FileName:="PathToFile" 
End Sub
```

В этом примере добавляется горизонтальную и вертикальную на активную страницу, которая пересекаются в центре страницы.




```vb
Sub SetRulerGuidesOnActivePage() 
 Dim intHeight As Integer 
 Dim intWidth As Integer 
 
 With ActiveView.ActivePage 
 intHeight = .Height / 2 
 intWidth = .Width / 2 
 With .RulerGuides 
 .Add Position:=intHeight, Type:=pbRulerGuideTypeHorizontal 
 .Add Position:=intWidth, Type:=pbRulerGuideTypeVertical 
 End With 
 End With 
End Sub
```


