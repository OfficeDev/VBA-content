---
title: "Свойство Font.DiacriticColor (издатель)"
keywords: vbapb10.chm5374003
f1_keywords: vbapb10.chm5374003
ms.prod: publisher
api_name: Publisher.Font.DiacriticColor
ms.assetid: 6e9c816e-c7ae-c559-6b35-150a5abb820c
ms.date: 06/08/2017
ms.openlocfilehash: 2285156f4312872ca10a0649d00e14ce514e86d3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontdiacriticcolor-property-publisher"></a>Свойство Font.DiacriticColor (издатель)

Возвращает объект **[ColorFormat](colorformat-object-publisher.md)** , представляющий 24-разрядный цвет, используемый для диакритические знаки в публикации языка для письма справа налево.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DiacriticColor**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

ColorFormat


## <a name="example"></a>Пример

В этом примере проверяется текст в первой статьи текущей публикации ли красный цвет и форматирования для письма справа налево.


```vb
Sub FontDiColor() 
 
 Dim fntDiColor As Font 
 
 Set fntDiColor = Application.ActiveDocument. _ 
 Stories(1).TextRange.Font 
 
 If fntDiColor.UseDiacriticColor = msoTrue And _ 
 fntDiColor.DiacriticColor.RGB = RGB(255, 0, 0) Then 
 MsgBox "Your text is red" 
 Else 
 MsgBox "This is not a right-to-left language" _ 
 &; " or your color is not red" 
 End If 
 
End Sub
```


