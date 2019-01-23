$file="C:\Users\Daniel\PycharmProjects\excal_test\xmlRead\resources\UI104635\LGV9B45B.xlss"
$xl=New-Object -ComObject "Excel.Application"
$xl.DisplayAlerts = $FALSE

$wb=$xl.Workbooks.Open($file)

$cells=$wb.ActiveSheet.Cells
$rows = $wb.ActiveSheet.UsedRange.Rows.Count

$wb.SaveAs("C:\Users\Daniel\PycharmProjects\excal_test\xmlRead\tmp\LGV9B45B.xls")

$wb.Close()
$xl.Quit()

while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)){}
while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)){}

python C:\Users\Daniel\PycharmProjects\excal_test\xmlRead\xmlEmail.py