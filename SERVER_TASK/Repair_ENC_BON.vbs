

'stvoriti excel objekt 
	Set objExcel = CreateObject("Excel.Application") 

'pogledajte program i datoteku excel, postavite na false da sakrijete cijeli postupak 
	objExcel.Visible = False'True 

'otvorite excel datoteku (obavezno promijenite mjesto) .xls za 2003. godinu ili ranije 
	Set objWorkbook = objExcel.Workbooks.Open("C:\Users\bivancic\Desktop\SERVER_TASK\REPAIR_database\Repair_ENC_BON.xlsm")


'Modul i COD u modulu
objExcel.Run "Repair_Access.RepairBaze"



'spremite postojeæu excel datoteku. Koristite SaveAs kako biste ga spremili kao nešto drugo. 
	objWorkbook.Save

'close the workbook
	objWorkbook.Close 

'exit the excel program
	objExcel.Quit




'release objects
	Set objExcel = Nothing
	Set objWorkbook = Nothing