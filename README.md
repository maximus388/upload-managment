# Проект "Управление загрузкой"

## Описание
Проект "Управление загрузкой" - инструмент управления Google-таблицами, позволяющий автоматически переносить данные из одной Google-таблицы в другую. Цель данного инструмента - автоматизировать копирование массивов данных из множества Google-таблиц. Его преимущество перед ручным копированием данных заключается в том, что пользователю нет необходимости самостоятельно открывать документы и копировать данные из одной таблицы в другую.

Ссылка на Google-диск с файлами: https://drive.google.com/drive/folders/1xQQH1M-mf97FoKomNB5Dgd_LIqhcQ8gQ

## Принцип работы
Для работы с инструментом необходимы 2 файла:
1. Google-таблица с инструкцией для скрипта, показывающая, откуда и куда нужно скопировать массив данных. Для копирования есть ряд настроек, позволяющих более гибко настроить копирование данных, таких как фильтр, тип вставки, адрес ячейки для вставки данных и другие.
Ссылка: https://docs.google.com/spreadsheets/d/1x78-8URsts_gZDAJYVM2d_x_whbDB-AUKbZOoN9KM1A/edit#gid=0

2. Сам скрипт для запуска копирования на платформе Google Colaboratory. Пользователю необходимо запустить скрипт, предварительно выбрав нужные таблицы для копирования и выбрав соответствующие настройки.
Ссылка: https://colab.research.google.com/drive/1sn2YscRqmD4A9AGYS3jABAZ2DICdmcpe#scrollTo=h_7qkG7O-M6X

## Порядок действий для работы с инструментом
1. Создать свою копию Google-таблицы с инструкцией для скрипта и копию самого скрипта в Google Colaboratory
2. Проверить наличие доступа ко всем участвующим в работе инструмента Google-таблицам
3. Заполнить данные по необходимым для обновления Google-таблицам в соответствии с примером на листе "Таблица 1". В заголовках параметров есть их описание и принцип работы 
4. В файле со скриптом необходимо заменить ссылку в значении переменной LINK_MAIN на ссылку созданной копии Google-таблицы с инструкцией для скрипта
5. Запустить скрипт, нажав на кнопку ⏵

## Планы на будущие обновления
1. Автоматическое копирование данных по расписанию
