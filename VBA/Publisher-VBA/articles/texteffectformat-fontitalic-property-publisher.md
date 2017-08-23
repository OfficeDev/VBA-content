---
title: "Свойство TextEffectFormat.FontItalic (издатель)"
keywords: vbapb10.chm3735810
f1_keywords: vbapb10.chm3735810
ms.prod: publisher
api_name: Publisher.TextEffectFormat.FontItalic
ms.assetid: 6594e6f7-e29e-a51d-55b8-d02f1fb9f26a
ms.date: 06/08/2017
ms.openlocfilehash: 953e734604b96c89c34b920ea8307c8e01992b65
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformatfontitalic-property-publisher"></a>Свойство TextEffectFormat.FontItalic (издатель)

Задает или возвращает константу **MsoTriState**, которое указывает, находится ли курсивное начертание шрифта для буквицы или текст надписи WordArt. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FontItalic**

 переменная _expression_A, представляет собой объект- **TextEffectFormat** .


## <a name="remarks"></a>Заметки

Значение свойства **FontItalic** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере вносятся буквицы в элементе frame указанного текста курсивом. В этом примере предполагается, что указанный текст frame отформатирован буквицы.


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


