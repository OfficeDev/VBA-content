---
title: "Свойство TextEffectFormat.Text (издатель)"
keywords: vbapb10.chm3735824
f1_keywords: vbapb10.chm3735824
ms.prod: publisher
api_name: Publisher.TextEffectFormat.Text
ms.assetid: eae1e95f-b0e6-559b-39a5-40291e758915
ms.date: 06/08/2017
ms.openlocfilehash: 513e652960c0dd2f0735a8c508767abb5cb4c605
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformattext-property-publisher"></a>Свойство TextEffectFormat.Text (издатель)

Возвращает или задает **строку** , представляющую текст в диапазон текста или WordArt фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Текст**

 переменная _expression_A, представляет собой объект- **TextEffectFormat** .


## <a name="example"></a>Пример

В следующем примере изменяется ее текст и задает имя шрифта и свойства форматирования для фигуры одно на первой странице active публикации. В данном примере для работы фигуры один должен быть объектом WordArt.


```vb
Sub FormatWordArt() 
 With ActiveDocument.Pages(1).Shapes(1).TextEffect 
 .Text = "This is a test." 
 .FontName = "Courier New" 
 .FontBold = True 
 .FontItalic = True 
 End With 
End Sub
```


