$file="C:\Users\da390967\PycharmProjects\xmlRead\xmlRead\resources\UI104635\LGV9B45B.xls"
$xl=New-Object -ComObject "Excel.Application"
$xl.DisplayAlerts = $FALSE

$wb=$xl.Workbooks.Open($file)

$cells=$wb.ActiveSheet.Cells
$rows = $wb.ActiveSheet.UsedRange.Rows.Count

$wb.SaveAs("C:\Users\da390967\PycharmProjects\xmlRead\xmlRead\tmp\LGV9B45B.xls")

$wb.Close()
$xl.Quit()

while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)){}
while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)){}

python --version