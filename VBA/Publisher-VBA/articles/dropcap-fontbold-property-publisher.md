---
title: "Свойство DropCap.FontBold (издатель)"
keywords: vbapb10.chm5505029
f1_keywords: vbapb10.chm5505029
ms.prod: publisher
api_name: Publisher.DropCap.FontBold
ms.assetid: 7e1b9b51-258d-080c-e5ae-cdc9d6a2ba64
ms.date: 06/08/2017
ms.openlocfilehash: a02f59d39a1892b10c681b154005925ec841243d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="dropcapfontbold-property-publisher"></a>Свойство DropCap.FontBold (издатель)

Задает или возвращает константу **MsoTriState**, которое указывает, находится ли шрифт для буквицы или влияние WordArt текст полужирным шрифтом. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FontBold**

 переменная _expression_A, представляет собой объект- **буквицу** .


## <a name="remarks"></a>Заметки

Значение свойства **FontBold** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере применяет жирное форматирование для буквицы в элементе frame указанный текст. В этом примере предполагается, что указанный текст frame отформатирован буквицы.


```vb
Sub BoldDropCap() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.DropCap 
 .FontBold = msoTrue 
 .FontColor.RGB = RGB(Red:=150, Green:=50, Blue:=180) 
 .FontItalic = msoTrue 
 .FontName = "Script MT Bold" 
 End With 
End Sub
```


