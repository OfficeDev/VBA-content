---
title: "Свойство AdvancedPrintOptions.IsPostscriptPrinter (издатель)"
keywords: vbapb10.chm7077921
f1_keywords: vbapb10.chm7077921
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.IsPostscriptPrinter
ms.assetid: 69c31e55-2781-38fa-7c4a-c5bc1b49972a
ms.date: 06/08/2017
ms.openlocfilehash: 6abe1755af26414de75c74793dc088111d2cd822
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsispostscriptprinter-property-publisher"></a>Свойство AdvancedPrintOptions.IsPostscriptPrinter (издатель)

Возвращает **значение True** , если active принтер PostScript принтера. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsPostscriptPrinter**

 переменная _expression_A, представляющий объект **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Следующие свойства объекта **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** доступны, только если active принтер Postscript принтера: **[HorizontalFlip](advancedprintoptions-horizontalflip-property-publisher.md)**, **[VerticalFlip](advancedprintoptions-verticalflip-property-publisher.md)**и **[NegativeImage](advancedprintoptions-negativeimage-property-publisher.md)**.

Свойство **[IsActivePrinter](printer-isactiveprinter-property-publisher.md)** используется для указания активного принтера для публикации.


## <a name="example"></a>Пример

Следующий пример определяет, является ли активного принтера PostScript принтера. Если он установлен, active публикация предназначена для печати, как зеркальное копирование по горизонтали и по вертикали, отрицательные изображение самого себя.


```vb
Sub PrepToPrintToFilmOnImagesetter() 
 
With ActiveDocument.AdvancedPrintOptions 
 If .IsPostscriptPrinter = True Then 
 .HorizontalFlip = True 
 .VerticalFlip = True 
 .NegativeImage = True 
 End If 
End With 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект AdvancedPrintOptions](advancedprintoptions-object-publisher.md)

