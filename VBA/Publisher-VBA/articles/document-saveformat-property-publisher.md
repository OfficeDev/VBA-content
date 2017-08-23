---
title: "Свойство Document.SaveFormat (издатель)"
keywords: vbapb10.chm196656
f1_keywords: vbapb10.chm196656
ms.prod: publisher
api_name: Publisher.Document.SaveFormat
ms.assetid: 545f0411-899f-ffe3-e844-8c2922a357f0
ms.date: 06/08/2017
ms.openlocfilehash: ff5f126c38b0cef1399476616997f736651107bb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentsaveformat-property-publisher"></a>Свойство Document.SaveFormat (издатель)

Указывает формат файла, указанного документа. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SaveFormat**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

PbFileFormat


## <a name="remarks"></a>Заметки

Значение свойства **SaveFormat** может иметь одно из **[PbFileFormat](pbfileformat-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

Если active публикации в формате Publisher 2000, в этом примере сохраняется в форматированный текст (RTF).


```vb
Sub SaveAsRTF() 
 
 If Application.ActiveDocument.SaveFormat = pbFilePublisher2000 Then 
 ActiveDocument.SaveAs "Flyer3", pbFileRTF 
 End If 
 
End Sub
```


