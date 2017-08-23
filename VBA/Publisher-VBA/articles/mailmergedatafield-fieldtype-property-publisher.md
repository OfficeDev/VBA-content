---
title: "Свойство MailMergeDataField.FieldType (издатель)"
ms.prod: publisher
api_name: Publisher.Field.FieldType
ms.assetid: 9574f59b-a03f-ab0b-a2ac-085f31473f78
ms.date: 06/08/2017
ms.openlocfilehash: 62a7017c4a34238b40d0d2a9a99bf36bab253fc3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafieldfieldtype-property-publisher"></a>Свойство MailMergeDataField.FieldType (издатель)

Возвращает константу **pbMailMergeDataFieldType** , представляющий тип данных, содержащихся в поле данных.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FieldType**

 переменная _expression_A, представляет собой объект- **MailMergeDataField** .


### <a name="return-value"></a>Возвращаемое значение

 **PbMailMergeDataFieldType**


## <a name="return-value"></a>Возвращаемое значение

 **PBMAILMERGEDATAFIELDTYPE**


## <a name="remarks"></a>Заметки

Используйте метод **[вставки](mailmergedatafield-insert-method-publisher.md)** объекта **[MailMergeDataField](mailmergedatafield-object-publisher.md)** Добавление поля данных изображения в области публикации.

Используйте метод **[InsertMailMergeField](textrange-insertmailmergefield-method-publisher.md)** объекта **[TextRange](textrange-object-publisher.md)** Добавление текстового поля данных в текстовом поле в области публикации.

Значение свойства **FieldType** может иметь одно из **[PbMailMergeDataFieldType](pbmailmergedatafieldtype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере определяет поля данных как поле данных изображения, вставляется в область данных для указанной публикации и размеры и располагает полей данных рисунка. В этом примере предполагается, что публикация подключен к источнику данных и что область объединения в каталог был добавлен к публикации.


```vb
Dim pbPictureField1 As Shape 
 
 'Define the Photo field as a picture data type 
 With ThisDocument.MailMerge.DataSource.DataFields 
 .Item("Photo:").FieldType = pbMailMergeDataFieldPicture 
 End With 
 
 'Insert a picture field, then size and position it 
 Set pbPictureField1 = ThisDocument.MailMerge.DataSource.DataFields.Item("Photo:").Insert 
 With pbPictureField1 
 .Height = 100 
 .Width = 100 
 .Top = 85 
 .Left = 375 
 End With
```


