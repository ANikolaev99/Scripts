#:: Этот блок обьявляет переменные
$date = get-date -uformat "%Y.%m.%d"
$d1 = "C:\Program Files\INFRATEL\integrations\Call-distribution-platform\CDP_*"
$d2 = "C:\Program Files\INFRATEL\integrations\Call-distribution-platform\Temp\Logs\"
$result = "C:\Program Files\INFRATEL\integrations\Call-distribution-platform\result.txt"

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
$To = "m.samokrutov@cdek.ru"
$Toto = "r.galikberova@cdek.ru"
$Totree = "nikalayeua@salestelecom.by"

#Тема письма
$subject = $date

#Логин и пароль от ящики с которого отправляете login@yandex.ru
$user = "itsupport@salestelecom.by"
$pass = "DobroeUtro1!"

#Путь до файла 
$file = "c:\Program Files\INFRATEL\integrations\Call-distribution-platform\Temp\Logs\$date.txt"

#Создаем два экземпляра класса
$att = New-object Net.Mail.Attachment($file)
$mes = New-Object System.Net.Mail.MailMessage

#Формируем данные для отправки
$mes.From = $from
$mes.To.Add($to) 
$mes.To.Add($Toto)
$mes.To.Add($Totree)
$mes.Subject = $subject 
$mes.IsBodyHTML = $true 
$mes.Body = "<h1>Log за прошлый день</h1>"

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