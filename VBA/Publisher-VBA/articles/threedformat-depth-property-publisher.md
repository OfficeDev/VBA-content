---
title: "Свойство ThreeDFormat.Depth (издатель)"
keywords: vbapb10.chm3801344
f1_keywords: vbapb10.chm3801344
ms.prod: publisher
api_name: Publisher.ThreeDFormat.Depth
ms.assetid: b6b46ddb-e3dd-0f9a-1a67-6433bb9ea89a
ms.date: 06/08/2017
ms.openlocfilehash: 9cd8b0a92b91c004b6273e394f39ff06162636f7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatdepth-property-publisher"></a>Свойство ThreeDFormat.Depth (издатель)

Возвращает или задает **Variant** , показывающее глубину придания объема фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Число уровней**

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Положительные значения необходимо создать придания объема, чьи лицевой является исходной фигуры; отрицательные значения осуществлять придания объема, чьи обратная фрагмент – исходной фигуры. Допустимые значения — от-600 через 9600 точек или эквивалентный расстояние в других единицах.


## <a name="example"></a>Пример

В этом примере добавляется овала active публикации и затем указывает, что овала быть вытянутый глубина 50 точек и выбирать должен быть фиолетовым.


```vb
Dim shpNew As Shape 
 
Set shpNew = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=90, Top:=90, Width:=90, Height:=40) 
 
With shpNew.ThreeD 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
End With 

```


