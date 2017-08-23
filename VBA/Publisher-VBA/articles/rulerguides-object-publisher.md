---
title: "Объект RulerGuides (издатель)"
keywords: vbapb10.chm786431
f1_keywords: vbapb10.chm786431
ms.prod: publisher
api_name: Publisher.RulerGuides
ms.assetid: c58d3cb2-8cf8-74fa-2bf4-a931dc95a26a
ms.date: 06/08/2017
ms.openlocfilehash: a3e45d73e3eb4c1239f542bed6c46646c04cdc24
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="rulerguides-object-publisher"></a>Объект RulerGuides (издатель)

Коллекция объектов **[RulerGuide](rulerguide-object-publisher.md)** , представляющий линий сетки, используется для выравнивания объектов на странице.
 


## <a name="example"></a>Пример

Используйте метод **[Add](rulerguides-add-method-publisher.md)** коллекции **RulerGuides** для добавления в коллекцию **RulerGuides** линии сетки линейки. В этом примере создается руководства по горизонтальной линейки и вертикальной направляющие каждый половины дюйм на первой странице active публикации.
 

 

```
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

Свойство **[Count](rulerguides-count-property-publisher.md)** возвращает общее число направляющие, горизонтальных и вертикальных в коллекции. В следующем примере свойство **Count** используйте цикл в котором удаление всех направляющие в коллекции.
 

 



```
Sub RemoveAllGuides() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).RulerGuides 
 For intCount = 1 To .Count 
 .Item(1).Delete 
 Next intCount 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](rulerguides-add-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](rulerguides-application-property-publisher.md)|
|[Count](rulerguides-count-property-publisher.md)|
|[Элемент](rulerguides-item-property-publisher.md)|
|[Родительский раздел](rulerguides-parent-property-publisher.md)|

