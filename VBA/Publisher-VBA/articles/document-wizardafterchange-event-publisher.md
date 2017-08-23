---
title: "Событие Document.WizardAfterChange (издатель)"
keywords: vbapb10.chm285212676
f1_keywords: vbapb10.chm285212676
ms.prod: publisher
api_name: Publisher.Document.WizardAfterChange
ms.assetid: c4ec0950-3a58-1f29-b35f-35db9d87f330
ms.date: 06/08/2017
ms.openlocfilehash: 7ef75316b6a6847d3f2cf29ee20fff5902a16469
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentwizardafterchange-event-publisher"></a>Событие Document.WizardAfterChange (издатель)

Возникает после пользователь выбирает вариант в окне мастера, которое изменяет любой из следующих параметров в публикации: макет страницы (размер страницы, сгиб тип, ориентация, метки продукта), Настройка печати (размер бумаги, размещение на странице), добавления и удаления объектов, добавление или удаление страниц, или объекта или страницы форматирование (размер, положение, заливки, границы, фон, текст по умолчанию, форматирование текста).


## <a name="syntax"></a>Синтаксис

 _выражение_. **WizardAfterChange**

 переменная _expression_A, представляющий объект **Document** .


## <a name="remarks"></a>Заметки

Событие WizardAfterChange возникает только один раз независимо от того, область или количество отдельных изменения, внесенные в публикации.

Для доступа к события объекта **Document** , объявите переменную объекта **документов** в разделе Общие описаний модуля класса, а затем задайте переменную равно объект **документа** , для которого требуется получить доступ к событиям.

Дополнительные сведения об использовании событий с помощью объекта **Document** содержатся в разделе [С помощью событий с помощью объекта Document](using-events-with-the-document-object-publisher.md).


## <a name="example"></a>Пример

В этом примере выводится сообщение при изменении с помощью панели мастера публикации. (Процедуры могут храниться в модуле ThisDocument публикации.)


```vb
Private Sub Document_WizardAfterChange() 
 MsgBox "Remember to save changes made " _ 
 &; "through the wizard pane!" 
End Sub
```


