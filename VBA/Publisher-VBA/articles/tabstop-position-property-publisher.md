---
title: "Свойство TabStop.Position (издатель)"
keywords: vbapb10.chm5636099
f1_keywords: vbapb10.chm5636099
ms.prod: publisher
api_name: Publisher.TabStop.Position
ms.assetid: 1ca7831a-6662-036e-8ba2-5784bc95fe8d
ms.date: 06/08/2017
ms.openlocfilehash: e36bc50a3276f0de9c9741ca2e9b8690804a7c3d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tabstopposition-property-publisher"></a>Свойство TabStop.Position (издатель)

Возвращает или задает **Variant** представляющее положение шрифта относительно базового плана текста в указанном диапазоне. Положительные значения переместить текст выше обычного базового, отрицательные значения переместить текст ниже базового плана. Неопределенное значения возвращаются в виде-9999.0. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Положение**

 переменная _expression_A, представляет собой объект- **TabStop** .


## <a name="example"></a>Пример

Этот пример устанавливает текст во второй материал на 5 точек ниже базового плана.


```vb
Sub Position() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Position = -5 
 
End Sub
```


