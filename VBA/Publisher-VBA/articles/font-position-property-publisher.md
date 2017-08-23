---
title: "Свойство Font.Position (издатель)"
keywords: vbapb10.chm5373988
f1_keywords: vbapb10.chm5373988
ms.prod: publisher
api_name: Publisher.Font.Position
ms.assetid: 24573faf-1627-3b10-5a8e-2f76a9f8831d
ms.date: 06/08/2017
ms.openlocfilehash: 2b2e319fd34b412d5787102547046d25f8e90a8d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontposition-property-publisher"></a>Свойство Font.Position (издатель)

Возвращает или задает **Variant** представляющее положение шрифта относительно базового плана текста в указанном диапазоне. Положительные значения переместить текст выше обычного базового, отрицательные значения переместить текст ниже базового плана. Неопределенное значения возвращаются в виде-9999.0. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Положение**

 переменная _expression_A, представляющий объект **Font** .


## <a name="example"></a>Пример

Этот пример устанавливает текст во второй материал на 5 точек ниже базового плана.


```vb
Sub Position() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Position = -5 
 
End Sub
```


