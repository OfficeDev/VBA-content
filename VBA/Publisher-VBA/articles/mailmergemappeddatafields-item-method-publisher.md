---
title: "Метод MailMergeMappedDataFields.Item (издатель)"
keywords: vbapb10.chm6488064
f1_keywords: vbapb10.chm6488064
ms.prod: publisher
api_name: Publisher.MailMergeMappedDataFields.Item
ms.assetid: c1c9acde-d1e5-25d3-1b59-3e848f3881b6
ms.date: 06/08/2017
ms.openlocfilehash: 320db070773a06b5381206e8082490b5f11649c5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergemappeddatafieldsitem-method-publisher"></a>Метод MailMergeMappedDataFields.Item (издатель)

Возвращает объект отдельных в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **MailMergeMappedDataFields** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Variant**|Номер или имя поля или поля элемента списка, чтобы возвратить.|

### <a name="return-value"></a>Возвращаемое значение

MailMergeMappedDataField


## <a name="example"></a>Пример

В этом примере возвращает поле «Город» из объекта поля сопоставленные данные.


```vb
Dim mmfTemp As MailMergeMappedDataField 
 
Set mmfTemp = ActiveDocument.MailMerge _ 
 .DataSource.MappedDataFields.Item(Index:="City")
```


