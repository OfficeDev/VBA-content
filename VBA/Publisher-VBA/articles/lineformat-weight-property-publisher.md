---
title: "Свойство LineFormat.Weight (издатель)"
keywords: vbapb10.chm3408147
f1_keywords: vbapb10.chm3408147
ms.prod: publisher
api_name: Publisher.LineFormat.Weight
ms.assetid: 854928ca-5f38-3cc9-50d5-2473a0885a0c
ms.date: 06/08/2017
ms.openlocfilehash: a0469af2ad4f3bac0fd958c92ab964d00e1fe6cf
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatweight-property-publisher"></a>Свойство LineFormat.Weight (издатель)

Возвращает или задает **Variant** , указывающее, толщины границы указанной строки или ячейки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вес**

 переменная _expression_A, представляет собой объект- **LineFormat** .


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


