---
title: "Метод ColorCMYK.SetCMYK (издатель)"
keywords: vbapb10.chm2621447
f1_keywords: vbapb10.chm2621447
ms.prod: publisher
api_name: Publisher.ColorCMYK.SetCMYK
ms.assetid: 9c7ec18b-73e9-66bc-57f4-cd6d62817630
ms.date: 06/08/2017
ms.openlocfilehash: 07c070863acbcd192998979f6405a2cdce1828ae
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorcmyksetcmyk-method-publisher"></a>Метод ColorCMYK.SetCMYK (издатель)

Задает значение голубой пурпурный желтый черный цвет (CMYK).


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetCMYK** ( **_Голубой_**, **_пурпурный_**, **_желтый_**, **_черный_**)

 переменная _expression_A, представляет собой объект- **ColorCMYK** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Голубой|Обязательное свойство.| **Длинный**|Число, представляющее голубой компонента цвета. Значение может быть любое число в диапазоне от 0 до 255.|
|Пурпурный|Обязательное свойство.| **Длинный**|Число, представляющее пурпурный компонента цвета. Значение может быть любое число в диапазоне от 0 до 255.|
|Желтый|Обязательное свойство.| **Длинный**|Число, представляющее компонент желтый цвет. Значение может быть любое число в диапазоне от 0 до 255.|
|Черный|Обязательное свойство.| **Длинный**|Число, представляющее компонент черного цвета. Значение может быть любое число в диапазоне от 0 до 255.|

## <a name="example"></a>Пример

В этом примере задается цвет CMYK для указанной фигуры.


```vb
Sub SetCMYKColor() 
 Dim shpStar As Shape 
 
 Set shpStar = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShape5pointStar, Left:=72, _ 
 Top:=72, Width:=150, Height:=150) 
 shpStar.Fill.ForeColor.CMYK.SetCMYK Cyan:=0, _ 
 Magenta:=255, Yellow:=255, Black:=50 
End Sub
```


