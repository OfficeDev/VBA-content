---
title: "Свойство AdvancedPrintOptions.VerticalFlip (издатель)"
keywords: vbapb10.chm7077891
f1_keywords: vbapb10.chm7077891
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.VerticalFlip
ms.assetid: d141d8c0-51a2-d47f-dda3-0cf273578b06
ms.date: 06/08/2017
ms.openlocfilehash: f88634f020d8cc65c5da9fe30eb9675043f5d51a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsverticalflip-property-publisher"></a>Свойство AdvancedPrintOptions.VerticalFlip (издатель)

 **Значение true** для печати по вертикали зеркальную указанной публикации. Значение по умолчанию — **False**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalFlip**

 переменная _expression_A, представляющий объект **AdvancedPrintOptions** .


## <a name="remarks"></a>Заметки

Это свойство доступно только в том случае, если active печати используется принтер PostScript. Возвращает ошибку времени выполнения, если указан другие принтера. Используйте свойство **[IsPostscriptPrinter](advancedprintoptions-ispostscriptprinter-property-publisher.md)** объекта **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** для определения, если указанный принтер PostScript принтера.

Данное свойство применяется для будущего экземпляры Microsoft Publisher и сохраняются как параметр приложения.

Это свойство соответствует управления **Отразить сверху вниз** на вкладке **Параметры страницы** диалоговое окно **Дополнительные параметры печати** .

Это свойство используется в основном при печати фильма фотонаборным, чтобы правильного считывания изображения, когда на лицевой стороне фильма не работает (как при записи на форме press).


## <a name="example"></a>Пример

Следующий пример определяет, является ли активного принтера PostScript принтера. Если он установлен, активных публикация предназначена для печати как по горизонтали зеркальное копирование и по вертикали зеркальное копирование, отрицательные изображение самого себя.


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

