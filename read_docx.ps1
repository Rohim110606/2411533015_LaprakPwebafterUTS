$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open("D:\Tugas SMT 5\Praktikum Aplikasi Mobile\PRAKTIKUM 5.docx")
$text = $doc.Content.Text
$doc.Close()
$word.Quit()
Write-Output $text