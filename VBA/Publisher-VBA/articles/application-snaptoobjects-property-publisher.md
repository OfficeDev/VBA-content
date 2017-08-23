---
title: "Свойство Application.SnapToObjects (издатель)"
keywords: vbapb10.chm131111
f1_keywords: vbapb10.chm131111
ms.prod: publisher
api_name: Publisher.Application.SnapToObjects
ms.assetid: 84fcb808-bf3b-49f7-666e-915ac6b04a96
ms.date: 06/08/2017
ms.openlocfilehash: 1f73119d4e3e0283e828d587aec1df3c643906b1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationsnaptoobjects-property-publisher"></a>Свойство Application.SnapToObjects (издатель)

 **Значение true** для Microsoft Publisher объекты на странице выровнять другие объекты. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SnapToObjects**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере добавляется горизонтальных и вертикальных линейки руководства по каждой половины дюйма на первой странице и настраивает параметры для выравнивания объектов на странице, чтобы руководства по.


```vb
Sub SetSnapOptions() 
 Dim intCount As Integer 
 Dim intPos As Integer 
 With ActiveDocument.Pages(1).RulerGuides 
 For intCount = 1 To 16 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeVertical 
 Next 
 intPos = 0 
 For intCount = 1 To 21 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeHorizontal 
 Next 
 End With 
 With Application 
 .SnapToGuides = True 
 .SnapToObjects = True 
 End With 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

