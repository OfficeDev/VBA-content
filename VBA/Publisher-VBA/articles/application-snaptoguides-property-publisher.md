---
title: "Свойство Application.SnapToGuides (издатель)"
keywords: vbapb10.chm131110
f1_keywords: vbapb10.chm131110
ms.prod: publisher
api_name: Publisher.Application.SnapToGuides
ms.assetid: 09894c02-3193-cd14-ff55-45920e461af9
ms.date: 06/08/2017
ms.openlocfilehash: b0c8aeaae517cf2eb697178db767dfe0787ef4b5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationsnaptoguides-property-publisher"></a>Свойство Application.SnapToGuides (издатель)

 **Значение true** для Microsoft Publisher направляющие выравнивания объектов на странице в публикации. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SnapToGuides**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере добавляется горизонтальных и вертикальных линейки руководства по каждой половины дюйма на первой странице и затем задает параметры для выравнивания объектов на странице, чтобы руководства по.


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

