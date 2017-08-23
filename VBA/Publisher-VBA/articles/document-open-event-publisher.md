---
title: "Событие Document.Open (издатель)"
keywords: vbapb10.chm285212673
f1_keywords: vbapb10.chm285212673
ms.prod: publisher
api_name: Publisher.Document.Open
ms.assetid: 43108d1d-d101-8a07-943e-c9b8dbadcbfd
ms.date: 06/08/2017
ms.openlocfilehash: 318dca8cc640bc1c2e78bb36d15e25795ab97832
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentopen-event-publisher"></a>Событие Document.Open (издатель)

Возникает при открытии публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Открыть**

 переменная _expression_A, представляющий объект **Document** .


## <a name="remarks"></a>Заметки

Для доступа к события объекта **Document** , объявите переменную объекта **документов** в разделе Общие описаний модуля класса, а затем задайте переменную равно объект **документа** , для которого требуется получить доступ к событиям.

Дополнительные сведения об использовании событий с помощью объекта **Document** содержатся в разделе [С помощью событий с помощью объекта Document](using-events-with-the-document-object-publisher.md).


## <a name="example"></a>Пример

В этом примере выводится сообщение при открытии публикации. (Процедуры могут храниться в модуле **ThisDocument** публикации.)


```vb
Private Sub Document_Open() 
 MsgBox "This publication is copyrighted." 
End Sub
```


