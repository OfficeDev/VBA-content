---
title: "Метод Document.Undo (издатель)"
keywords: vbapb10.chm196704
f1_keywords: vbapb10.chm196704
ms.prod: publisher
api_name: Publisher.Document.Undo
ms.assetid: 8cfd09a0-8a0d-2870-f833-a35ff1fc21b4
ms.date: 06/08/2017
ms.openlocfilehash: eada9e30aa8734670714c5465715b6a89744060d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentundo-method-publisher"></a>Метод Document.Undo (издатель)

Отменяет последнее действие или указанное число действий. Соответствует список элементов, который отображается, если щелкнуть стрелку рядом с кнопкой " **Отмена** " на панели инструментов **Стандартная** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Отменить** ( **_Количество_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Count|Необязательный| **Длинный**|Задает число действий, чтобы быть отменена. Значение по умолчанию — 1, что означает, что если этот параметр опущен, отменить последнее действие.|

## <a name="remarks"></a>Заметки

Если вызывается, когда нет никаких действий в стеке отмены или **_Count_** больше, чем число действий, которые в настоящее время находятся в стеке, метод **Отменить** отмены всех действий по мере возможности без применения rest.

Максимальное число действий, которые можно отменить в одном вызове для **отмены** — 20.


## <a name="example"></a>Пример

В следующем примере метод **Отменить** для отмены всех действий, которые не отвечают определенным условиям.

Часть 1 из примера в четвертой странице активная публикация добавляется фигура прямоугольный выноски и текст добавляется в выноске. Этот процесс создает три действия. 

Часть 2 примера проверяет, является ли шрифт текста, добавляемого на выноске Verdana. В противном случае выберите метод **Отменить** используется для отмены всех доступных действий (значение свойства **[UndoActionsAvailable](document-undoactionsavailable-property-publisher.md)** используется для указания, что все действия быть отменены). Это приведет к очистке всех действий из стека. Затем добавляются новые прямоугольника фигуры и текст frame и выполняется заполнение рамки с текстом.




```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(4) 
 
With theDoc 
 ' Part 1 
 With thePage 
 ' Setting the shape creates the first action 
 Set theShape = .Shapes.AddShape(msoShapeRectangularCallout, _ 
 75, 75, 120, 30) 
 ' Setting the text range creates the second action 
 With theShape.TextFrame.TextRange 
 ' Setting the text creates the third action 
 .Text = "This text is not Verdana." 
 End With 
 End With 
 
 ' Part 2 
 If Not thePage.Shapes(1).TextFrame.TextRange.Font.Name = "Verdana" Then 
 ' UndoActionsAvailable = 3 
 .Undo (.UndoActionsAvailable) 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeRectangle, _ 
 75, 75, 120, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Name = "Verdana" 
 .Text = "This text is Verdana." 
 End With 
 End With 
 End If 
End With
```


