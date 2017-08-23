---
title: "Свойство Document.ReadOnly (издатель)"
keywords: vbapb10.chm196647
f1_keywords: vbapb10.chm196647
ms.prod: publisher
api_name: Publisher.Document.ReadOnly
ms.assetid: 9ee6488d-3070-e784-e772-78dace2c1284
ms.date: 06/08/2017
ms.openlocfilehash: 4a4cbc4712670100b8465c251127b2d40ec6cab9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentreadonly-property-publisher"></a>Свойство Document.ReadOnly (издатель)

Возвращает **значение True,** Если публикации только для чтения. Возвращает **значение False** , если это чтения и записи. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Только для чтения**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере сохраняет active публикации и уведомляет пользователя, которая сохраняется в файле и является ли оно только для чтения.


```vb
Sub SaveAndStatus() 
 
 Dim bStatus As Boolean 
 
 Application.ActiveDocument.SaveAs "c:\testfile.pub" 
 bStatus = Application.ActiveDocument.ReadOnly 
 MsgBox "File Saved and Read-only Status = " &; bStatus 
 
End Sub
```


