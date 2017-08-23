---
title: "Свойство Font.AutomaticPairKerningThreshold (издатель)"
keywords: vbapb10.chm5373975
f1_keywords: vbapb10.chm5373975
ms.prod: publisher
api_name: Publisher.Font.AutomaticPairKerningThreshold
ms.assetid: f5f43a19-7227-b25d-9322-84a79596c525
ms.date: 06/08/2017
ms.openlocfilehash: d2af9610b4b477427614c537f3c2f3effd8e190c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontautomaticpairkerningthreshold-property-publisher"></a>Свойство Font.AutomaticPairKerningThreshold (издатель)

Возвращает или задает значение **типа Variant** , представляющее размер шрифта, над которой кернинг автоматически настраивается для символов в диапазоне указанный текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AutomaticPairKerningThreshold**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Допустимые значения — от 0,0 указывает на 999,5 пунктов. Возвращает -2, если значение для символов в диапазоне текст не определено. Назначить этому свойству значение 0.0 отключает автоматическую пары кернинг диапазона.


## <a name="example"></a>Пример

В этом примере задается порог размер точки 12 пунктов. Весь текст во второй материал выше порога реализации кернинг автоматически.


```vb
Sub Threshold() 
 
 Application.ActiveDocument.Stories(2).TextRange _ 
 .Font.AutomaticPairKerningThreshold = 12 
 
End Sub
```


