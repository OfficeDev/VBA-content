---
title: "Метод Plate.ConvertToProcess (издатель)"
keywords: vbapb10.chm2883601
f1_keywords: vbapb10.chm2883601
ms.prod: publisher
api_name: Publisher.Plate.ConvertToProcess
ms.assetid: 26476701-aa82-ca44-20c8-55a332a6539a
ms.date: 06/08/2017
ms.openlocfilehash: e7a26bde34c0ca3b1187ab62719d102f92878b2f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="plateconverttoprocess-method-publisher"></a>Метод Plate.ConvertToProcess (издатель)

Преобразует указанной формы из плашечных для обработки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ConvertToProcess**

 переменная _expression_A, представляющий объект **формы** .


## <a name="remarks"></a>Заметки

Метод **ConvertToProcess** доступна только в том случае, если режим цвета публикации имеет значение процесс и плашечные цвета. Использование ** [EnterColorMode](http://msdn.microsoft.com/library/3c04275d-d274-f681-7391-139a54232a3b%28Office.15%29.aspx)** метод объекта **[Document](document-object-publisher.md)** , чтобы задать режим цвета публикации.

Возвращает «Отказано в разрешении» применительно к форме плашечных цветов. Если режим цвет включает процесс цвет, цветов процесс (черный, пурпурный, желтый и голубой) являются сначала четыре формы в коллекции **[формы](plates-object-publisher.md)** .

При преобразовании форму из место для обработки цвет, все цвета публикации на основании рукописного ввода, представляющий преобразованные формы преобразуются в обработки цвета.


## <a name="example"></a>Пример

В следующем примере преобразуется форме указанного плашечных цветов для обработки цвета. Предполагается, что режим цвета публикации был указан как место и процесс цвета и, по крайней мере шесть формы были определены для публикации.


```vb
Sub ChangePlateToProcess() 
 
 With ActiveDocument.Plates.Item(6) 
 .ConvertToProcess 
 End With 
 
End Sub
```


