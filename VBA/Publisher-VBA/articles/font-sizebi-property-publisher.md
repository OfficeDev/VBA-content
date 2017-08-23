---
title: "Свойство Font.SizeBi (издатель)"
keywords: vbapb10.chm5373958
f1_keywords: vbapb10.chm5373958
ms.prod: publisher
api_name: Publisher.Font.SizeBi
ms.assetid: 1e9100e7-efa4-a7aa-69af-39c550a0b046
ms.date: 06/08/2017
ms.openlocfilehash: 5697b307740c6a44e1b554ec752eb6c3191159cc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontsizebi-property-publisher"></a>Свойство Font.SizeBi (издатель)

Возвращает или задает значение **типа Variant** , представляющее размер в пунктах объекта **шрифта** для текста справа налево языке. Допустимые значения — от 0,5 указывает на 999,5 пунктов. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SizeBi**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="example"></a>Пример

В этом примере проверяется текст во второй материал. Если справа налево языке, размер которых превышает 12 точек и курсив, полужирный — это набор текст.


```vb
Sub SizeBiIfBig() 
 
 Dim fntSize As Font 
 
 Set fntSize = Application.ActiveDocument.Stories(2).TextRange.Font 
 With fntSize 
 If .SizeBi > 12 And .ItalicBi = msoTrue Then 
 .BoldBi = msoTrue 
 Else 
 MsgBox "The font size is 12 points or less " &; _ 
 ", not bold, or this is not in a right-to-left language." 
 End If 
 End With 
 
End Sub
```


