#:: Этот блок обьявляет переменные
$date = get-date -uformat "%Y.%m.%d"
$d1 = "C:\Works\Temp\scripts\log.txt"
$d2 = "C:\Works\Temp\scripts\Temp\"
$result = "C:\Works\Temp\scripts\result.txt"

#:: Этот блок записывает в файл результат выпонлнения
"Скопировано $date в $d2" >> $result

#:: Этот блок перемещает log файл
#move "%d1%" "%d2%%date%.txt"
Move-Item -Path $d1 -Destination $d2$date.txt

#Адрес сервера SMTP для отправки
$serverSmtp = "smtp.yandex.ru" 

#Порт сервера
$port = 587

#От кого
$From = "itsupport@salestelecom.by" 

#Кому
$To = "nikolka.99@gmail.com" 

#Тема письма
$subject = $date

#Логин и пароль от ящики с которого отправляете login@yandex.ru
$user = "itsupport@salestelecom.by"
$pass = "DobroeUtro1!"

#Путь до файла 
$file = "C:\Works\Temp\scripts\Temp\$date.txt"

#Создаем два экземпляра класса
$att = New-object Net.Mail.Attachment($file)
$mes = New-Object System.Net.Mail.MailMessage

#Формируем данные для отправки
$mes.From = $from
$mes.To.Add($to) 
$mes.Subject = $subject 
$mes.IsBodyHTML = $true 
$mes.Body = "<h1>Тестовое письмо</h1>"

#Добавляем файл
$mes.Attachments.Add($att) 

#Создаем экземпляр класса подключения к SMTP серверу 
$smtp = New-Object Net.Mail.SmtpClient($serverSmtp, $port)

#Сервер использует SSL 
$smtp.EnableSSL = $true 

#Создаем экземпляр класса для авторизации на сервере яндекса
$smtp.Credentials = New-Object System.Net.NetworkCredential($user, $pass);

#Отправляем письмо, освобождаем память
$smtp.Send($mes) 
$att.Dispose()