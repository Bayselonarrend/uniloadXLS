//MIT License

//Copyright (c) 2023 Anton Tsitavets

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

# Область Excel 

//!!! Пример вызова на форме  !!!

//&НаКлиенте
//Процедура ЗагрузитьXLS(Команда)
//	
//	СтруктураКолонок = Новый Структура;
//	СтруктураКолонок.Вставить("Спр2"			, "СправочникСсылка.Справочник2");
//	СтруктураКолонок.Вставить("Перечисление1"	        , "ПеречислениеСсылка.Перечисление1");
//	СтруктураКолонок.Вставить("ЛюбоеИмя"		        , "");
//	СтруктураКолонок.Вставить("Док1"			, "ДокументСсылка.Документ1");
//	СтруктураКолонок.Вставить("Спр1"			, "СправочникСсылка.Справочник1");

//	
//	ИД = КлиентСерверный.ЗагрузитьИзXLS(СтруктураКолонок,,2);
//	
//	Если ЗначениеЗаполнено(ИД)  Тогда
//		ЗагрузкаНаСервере(ИД);
//	КонецЕсли; 
//	
//КонецПроцедуры


//&НаСервере
//Процедура ЗагрузкаНаСервере(ИД)	
//	ТЗ = ПолучитьИзВременногоХранилища(ИД);
//КонецПроцедуры


&НаКлиенте
Функция ЗагрузитьИзXLS(СоответствиеКолонок, НазваниеЛиста = "", НомерПервойСтроки = 1) Экспорт 
	
	ИД 	= Серверный.ПоместитьЗаглушку(СоответствиеКолонок);
	
	ДиалогВыбораФайла 				= Новый ДиалогВыбораФайла(РежимДиалогаВыбораФайла.Открытие);
	ДиалогВыбораФайла.Фильтр 			= "Документ Excel (*.xls, *.xlsx)|*.xls;*.xlsx|"; 
	ДиалогВыбораФайла.Заголовок 			= "Выберите файл";
	ДиалогВыбораФайла.ПредварительныйПросмотр 	= Ложь;
	ДиалогВыбораФайла.МножественныйВыбор 		= Ложь;
	ДиалогВыбораФайла.ИндексФильтра 		= 0;
	
	Параметры 		= Новый Структура("Макет,АдресХр,НазваниеЛиста,НомерПервойСтроки", СоответствиеКолонок, ИД, НазваниеЛиста, НомерПервойСтроки);
	//ОписаниеОповещения 	= Новый ОписаниеОповещения("ЗагрузитьФайлЗавершение", ЭтотОбъект, Параметры);
	//НачатьПомещениеФайла(ОписаниеОповещения,,ДиалогВыбораФайла,Истина);
	
	Если ДиалогВыбораФайла.Выбрать() Тогда
		ПоместитьВоВременноеХранилище(Новый ДвоичныеДанные(ДиалогВыбораФайла.ПолноеИмяФайла), ИД);
		ЗагрузитьФайлЗавершение(Истина, Ид, ДиалогВыбораФайла.ПолноеИмяФайла, Параметры);
	КонецЕсли;
	
	Возврат ИД;
	
КонецФункции

Процедура ЗагрузитьФайлЗавершение(Результат, Адрес, ВыбранноеИмяФайла, ДополнительныеПараметры) Экспорт
	
	ИмяФайла = Прав(ВыбранноеИмяФайла, СтрДлина(ВыбранноеИмяФайла) - СтрНайти(ВыбранноеИмяФайла, "\", НаправлениеПоиска.СКонца));
						
	Если Не Адрес = "" Тогда
		Серверный.ФормированиеТаблицы(Адрес, ИмяФайла, ДополнительныеПараметры);
	КонецЕсли;
	
КонецПроцедуры

# КонецОбласти 
