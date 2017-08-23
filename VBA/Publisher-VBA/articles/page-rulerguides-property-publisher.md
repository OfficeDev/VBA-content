---
title: "Свойство Page.RulerGuides (издатель)"
keywords: vbapb10.chm393225
f1_keywords: vbapb10.chm393225
ms.prod: publisher
api_name: Publisher.Page.RulerGuides
ms.assetid: 69605642-7722-0721-cb07-d33689eda9ab
ms.date: 06/08/2017
ms.openlocfilehash: edd52404209acd20aebac89a10d2234720615978
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagerulerguides-property-publisher"></a>Свойство Page.RulerGuides (издатель)

Возвращает коллекцию **[RulerGuides](rulerguides-object-publisher.md)** , представляющий линии сетки используется для выравнивания объектов на странице.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RulerGuides**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

RulerGuides


## <a name="example"></a>Пример

В этом примере создается руководства по горизонтальной линейки и вертикальной направляющие каждый половины дюйм на первой странице active публикации.


```vb
Sub SetRulerGuides() 
 Dim intCount As Integer 
 Dim intPos As Integer 
 With ActiveDocument.Pages(1).RulerGuides 
 For intCount = 1 To 16 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeVertical 
 Next intCount 
 intPos = 0 
 For intCount = 1 To 21 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeHorizontal 
 Next intCount 
 End With 
End Sub
```


