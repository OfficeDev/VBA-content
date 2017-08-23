---
title: "Метод MailMergeDataField.Insert (издатель)"
keywords: vbapb10.chm6422561
f1_keywords: vbapb10.chm6422561
ms.prod: publisher
api_name: Publisher.MailMergeDataField.Insert
ms.assetid: 54482cda-d0d3-c799-7e7f-b25835a8bd6f
ms.date: 06/08/2017
ms.openlocfilehash: 1939c327ec170dca3586d26c6960b798f315de2f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafieldinsert-method-publisher"></a>Метод MailMergeDataField.Insert (издатель)

Возвращает объект **[фигуры](shape-object-publisher.md)** , который представляет поле данных, вставленных в публикацию.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вставка** ( **_Диапазон_**)

 переменная _expression_A, представляет собой объект- **MailMergeDataField** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Range|Необязательный| **TextRange**|Диапазон текста для вставки.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

**Вставка** работает метод для обоих рисунок и строковые поля (текст).


 **Примечание**  Можно также использовать метод **[InsertMailMergeField](textrange-insertmailmergefield-method-publisher.md)** объекта **[TextRange](textrange-object-publisher.md)** для добавления текстового поля данных в текстовом поле в области публикации.


## <a name="example"></a>Пример

В этом примере определяет поля данных как поле данных изображения, вставляется в область данных для указанной публикации и размеры и располагает полей данных рисунка. В этом примере предполагается публикации подключен к источнику данных, а область объединения в каталог был добавлен к публикации.


```vb
Dim pbPictureField1 As Shape 
 
 'Define the field as a picture data type 
 With ThisDocument.MailMerge.DataSource.DataFields 
 .Item("Photo:").FieldType = pbMailMergeDataFieldPicture 
 End With 
 
 'Insert a picture field, and then size and position it 
 Set pbPictureField1 = ThisDocument.MailMerge.DataSource.DataFields.Item("Photo:").Insert 
 With pbPictureField1 
 .Height = 100 
 .Width = 100 
 .Top = 85 
 .Left = 375 
 End With
```


