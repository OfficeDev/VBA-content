---
title: "Свойство RulerGuide.Position (издатель)"
keywords: vbapb10.chm655364
f1_keywords: vbapb10.chm655364
ms.prod: publisher
api_name: Publisher.RulerGuide.Position
ms.assetid: af169eaf-3cca-0310-c49b-369ba9b2193f
ms.date: 06/08/2017
ms.openlocfilehash: 6c8c58d80d55d3834caf4b441210c7cb357df56a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="rulerguideposition-property-publisher"></a>Свойство RulerGuide.Position (издатель)

Возвращает или задает **Variant** представляющее положение шрифта относительно базового плана текста в указанном диапазоне. Положительные значения переместить текст выше обычного базового, отрицательные значения переместить текст ниже базового плана. Неопределенное значения возвращаются в виде-9999.0. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Положение**

 переменная _expression_A, представляет собой объект- **RulerGuide** .


## <a name="example"></a>Пример

Этот пример устанавливает текст во второй материал на 5 точек ниже базового плана.


```vb
Sub Position() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Position = -5 
 
End Sub
```


