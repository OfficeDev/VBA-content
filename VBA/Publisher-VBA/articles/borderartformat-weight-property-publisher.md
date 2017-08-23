---
title: "Свойство BorderArtFormat.Weight (издатель)"
keywords: vbapb10.chm7602182
f1_keywords: vbapb10.chm7602182
ms.prod: publisher
api_name: Publisher.BorderArtFormat.Weight
ms.assetid: 8ff67c8b-be41-a02e-5433-624baa0d888e
ms.date: 06/08/2017
ms.openlocfilehash: 9d891bd65f6e95c292b81821a0fd21b8a91b2810
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartformatweight-property-publisher"></a>Свойство BorderArtFormat.Weight (издатель)

Возвращает или задает **Variant** , указывающее, толщины границы указанной строки или ячейки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вес**

 переменная _expression_A, представляет собой объект- **BorderArtFormat** .


## <a name="remarks"></a>Заметки

Возвращаемые значения находятся в пунктах. При задании свойства числовые значения вычисляются в точках и строк может быть в любой устройств, поддерживаемых Publisher (например, «2,5 дюйма»).


## <a name="example"></a>Пример

В этом примере добавляет зеленой пунктирной линии, два аспекта толстые, активных публикации.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddLine(BeginX:=10, BeginY:=10, _ 
 EndX:=250, EndY:=250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(0, 255, 255) 
 .Weight = 2 
End With 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект BorderArtFormat](borderartformat-object-publisher.md)

