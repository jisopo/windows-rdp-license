Скрипт проводит активацию rdp лицензий для windows server 2012/2016 через веб-браузер с использованием selenium. На выходе получаем ключ для активации сервера + клиентов(пользователей/устройств).

## Зависимости(лежат в папке Redist):
* python 3.5.2
* chrome driver
* модули для pip

## Модули pip:
* pysmb

## Краткое описание как все вместе работает:
1) Ставим powershell модули для установки лицензий на windows server который будем активировать.
2) Получаем с сервера product_id.txt для активации с его помощью в веб-браузере.
3) Скрипт активирует через браузер лицензии с помощью ключа из product_id.txt и на выходе получаем 2 файла: server_activation_key.txt и client_license_activation_key.txt

server_activation_key.txt - для активации сервера лицензий

client_license_activation_key.txt - для активации лицензии на пользователей или устройства

## Подготовка:

**На windows сервере**

1) Устанавливаем компоненты через powershell

````
Install-WindowsFeature RDS-Licensing -IncludeAllSubFeature -IncludeManagementTools
````

2) Вытаскиваем id продукта и сохраняем на samba сервере file-server через powershell(не забываем поменять имя пользователя в пути)

````
(Get-WmiObject Win32_TSLicenseServer).ProductId | Out-file -Encoding ASCII -FilePath \\file-server\\incoming\\jisopo\\product_id.txt
````

**Настройки скрипта**

Первоначально скрипт регистрирует лицензии для 10 объектов и по умолчанию регистрирует лицензию для устройств(devices). Количество требуемых лицензий можно поменять через переменную PREFERED_LICENSE_COUNT. Для указания выдачи лицензий на пользователей нужно SELECT_ACTIVATION_BY_DEVICES установить в False. Дополнительные настройки, такие как название компании, страна и прочее можно также указать в настройках скрипта: PREFERED_COMPANY_NAME и PREFERED_COMPANY_COUNTRY соответственно.

Для активации лицензии на 2012 сервере поменять цифры с 2016 на 2012 в полях **PREREFED_CALS_LICENSE_TYPE_DEVICE** и **PREREFED_CALS_LICENSE_TYPE_USER**

## Использование:

**На windows машине с выходом в интернет**

3) Меняем настройки скрипта под свои нужны и запускаем auto_license.bat. На выходе получаем файлы server_activation_key.txt и client_license_activation_key.txt
