---
title: "Метод Document.Redo (издатель)"
keywords: vbapb10.chm196708
f1_keywords: vbapb10.chm196708
ms.prod: publisher
api_name: Publisher.Document.Redo
ms.assetid: 4b76aeaa-77f7-5f22-ff80-77479b0f0702
ms.date: 06/08/2017
ms.openlocfilehash: 44392f7aeb335f3d04104d1396ab04227cc2a6ba
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentredo-method-publisher"></a>Метод Document.Redo (издатель)

Повтор последнего действия или указанное число действий. Соответствует список элементов, который отображается, если щелкнуть стрелку рядом с кнопкой **Вернуть** на панели инструментов **Стандартная** . Вызывающий этот метод отменяет ** [Отменить метод](document-undo-method-publisher.md)** метод.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Повторное применение** ( **_Количество_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Count|Необязательный| **Длинный**|Задает число действий, должен быть перезаписан. Значение по умолчанию — 1, что означает, что если этот параметр опущен, будет перезаписан только последнее действие.|

### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Если вызывается, когда нет никаких действий в стеке повтора или **_Count_** больше, чем число действий, которые в настоящее время находятся в стеке, метод **Повтор** Повтор всех действий по мере возможности без применения rest.

Максимальное число действий, которые можно применить в одном вызове **Возврат** — 20.


## <a name="example"></a>Пример

В следующем примере метод **Повтор** Повтор подмножество действия, которые были отменены с помощью метода **Отменить** .

Часть 1 создает прямоугольник, который содержит фрагмент текста на четвертой странице active публикации. Задать различные свойства шрифта и добавлен текст надписи. В данном случае текст «этот шрифт является Courier» задано значение полужирным шрифтом Courier 12 пунктов. 

Часть 2 проверяет, является ли текст в текстовой рамке шрифт Verdana. В противном случае выберите метод **Отменить** используется для отмены всех последние четыре действия в стеке отмены. Метод **Повторить** нажмите используется для возврата первых двух последние четыре действия, которые были только что отменены. В этом случае третий действие (размер шрифта) и четвертый (Настройка полужирный шрифт) действия повторно выполняется. Имя шрифта нажмите Изменить для Verdana и изменить текст.




```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(4) 
 
' Part 1 
With theDoc 
 With thePage 
 ' Setting the shape creates the first action 
 Set theShape = .Shapes.AddShape(msoShapeRectangle, _ 
 75, 75, 190, 30) 
 ' Setting the text range creates the second action 
 With theShape.TextFrame.TextRange 
 ' Setting the font size creates the third action 
 .Font.Size = 12 
 ' Setting the font to bold creates the fourth action 
 .Font.Bold = msoTrue 
 ' Setting the font name creates the fifth action 
 .Font.Name = "Courier" 
 ' Setting the text creates the sixth action 
 .Text = "This font is Courier." 
 End With 
 End With 
 
 ' Part 2 
 If Not thePage.Shapes(1).TextFrame.TextRange.Font.Name = "Verdana" Then 
 .Undo (4) 
 With thePage 
 With theShape.TextFrame.TextRange 
 ' Redo redoes the first two of the four actions that were just undone 
 theDoc.Redo (2) 
 .Font.Name = "Verdana" 
 .Text = "This font is Verdana." 
 End With 
 End With 
 End If 
End With
```


