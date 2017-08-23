---
title: "Объект RulerGuide (издатель)"
keywords: vbapb10.chm720895
f1_keywords: vbapb10.chm720895
ms.prod: publisher
api_name: Publisher.RulerGuide
ms.assetid: 6400c368-02e9-169c-c675-9416cd361384
ms.date: 06/08/2017
ms.openlocfilehash: 6d3d4bc0b03846f7e7634c90426b9d6ac348e639
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="rulerguide-object-publisher"></a>Объект RulerGuide (издатель)

Представляет линий сетки, используется для выравнивания объектов на странице. Объект **RulerGuide** является элементом коллекции **[RulerGuides](rulerguides-object-publisher.md)** .
 


## <a name="example"></a>Пример

Используйте метод **[Add](rulerguides-add-method-publisher.md)** коллекции **RulerGuides** для создания новой линии сетки линейки. Используйте свойство **[Item](rulerguides-item-property-publisher.md)** для ссылки направляющей линейки. Используйте свойство **[положение](rulerguide-position-property-publisher.md)** изменение положения линии сетки и использование метода **[Delete](rulerguide-delete-method-publisher.md)** для удаления линий сетки. В этом примере создается новый направляющей линейки, переводит его, а затем удаляет его.
 

 

```
Sub AddChangeDeleteGuide() 
 Dim rgLine As RulerGuide 
 With ActiveDocument.Pages(1).RulerGuides 
 .Add Position:=InchesToPoints(1), _ 
 Type:=pbRulerGuideTypeVertical 
 
 MsgBox "The ruler guide position is at one inch." 
 
 .Item(1).Position = InchesToPoints(3) 
 MsgBox "The ruler guide is now at three inches." 
 
 .Item(1).Delete 
 MsgBox "The ruler guide has been deleted." 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](rulerguide-delete-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](rulerguide-application-property-publisher.md)|
|[Родительский раздел](rulerguide-parent-property-publisher.md)|
|[Position](rulerguide-position-property-publisher.md)|
|[Type](rulerguide-type-property-publisher.md)|

