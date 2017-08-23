---
title: "Метод MailMergeDataSources.Item (издатель)"
keywords: vbapb10.chm7143427
f1_keywords: vbapb10.chm7143427
ms.prod: publisher
api_name: Publisher.MailMergeDataSources.Item
ms.assetid: a65fedf6-aae5-64ef-e7d0-6bbc3d5b733c
ms.date: 06/08/2017
ms.openlocfilehash: 3497fe3397d535d24d06042d3589550c887c91a4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcesitem-method-publisher"></a>Метод MailMergeDataSources.Item (издатель)

Возвращает объект **[вывода](mailmergedatasource-object-publisher.md)** по указанному индексу в коллекции **MailMergeDataSources** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_varIndex_**)

 переменная _expression_A, представляющий коллекцию **MailMergeDataSources** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|varIndex|Обязательное свойство.| **Variant**|Номер индекса или имя возвращаемого объекта.|

### <a name="return-value"></a>Возвращаемое значение

Вывода


## <a name="remarks"></a>Заметки

Метод **Item** — это элемент по умолчанию коллекции **MailMergeDataSources**

Если имеется только один объект **вывода** в активном документе, коллекция **MailMergeDataSources** пуста. В этом случае при попытке использовать свойство **DataSources** объекта **вывода** для получения коллекции источников данных, Microsoft Publisher возвращает ошибку.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как получить имена всех подключенных источников данных в коллекции **MailMergeDataSources** в активный документ. Он используется свойство **IsDataSourceConnected** активного документа для определения, подключен ли в источнике данных.

Если подключение одного или нескольких источников данных, макрос использует свойство **Count** коллекции **MailMergeDataSources** для определения подключенные сколько источников данных.

Если подключен только один источник данных, макрос печатает имя этого источника данных в окне **интерпретации** ; Если подключено более одного источника данных, макрос использует метод **элемента** коллекции **MailMergeDataSources** для итерации по коллекции и свойство **Name** объекта **вывода** для печати имя каждого источника данных в окне **интерпретации** .




```vb
Public Sub Item_Example() 
 
 Dim pubMailMergeDataSources As Publisher.MailMergeDataSources 
 Dim pubMailMergeDataSource As Publisher.MailMergeDataSource 
 Dim lngCount As Long 
 Dim intCounter As Integer 
 
 If ThisDocument.IsDataSourceConnected Then 
 
 Set pubMailMergeDataSources = ThisDocument.MailMerge.DataSource.DataSources 
 
 lngCount = pubMailMergeDataSources.Count 
 
 If lngCount > 1 Then 
 
 ' More than one data source is connected. 
 For intCounter = 1 To lngCount 
 Debug.Print pubMailMergeDataSources.Item(intCounter).Name 
 Next 
 
 Else 
 
 ' Only one data source is connected. 
 Set pubMailMergeDataSource = ThisDocument.MailMerge.DataSource 
 Debug.Print "Only one data source ("; pubMailMergeDataSource.Name; ") is connected." 
 
 End If 
 
 Else 
 
 Debug.Print "No data sources are connected." 
 
 End If 
 
End Sub
```


