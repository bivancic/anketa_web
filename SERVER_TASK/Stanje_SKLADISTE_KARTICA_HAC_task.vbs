

'stvoriti excel objekt 
	Set objExcel = CreateObject("Excel.Application") 

'pogledajte program i datoteku excel, postavite na false da sakrijete cijeli postupak 
	objExcel.Visible = False'True 

'otvorite excel datoteku (obavezno promijenite mjesto) .xls za 2003. godinu ili ranije 
	Set objWorkbook = objExcel.Workbooks.Open("C:\Users\bivancic\Desktop\SERVER_TASK\STANJE_database\SKLADISTE_KARTICA_HAC_TASK_v1.xlsm")


'Modul i COD u modulu
objExcel.Run "Petlja_NA_SKLADISTU_SVI.Pokreni_Izlistaj_Sve_NP"



'spremite postojeæu excel datoteku. Koristite SaveAs kako biste ga spremili kao nešto drugo. 
	objWorkbook.Save

'close the workbook
	objWorkbook.Close 

'exit the excel program
	objExcel.Quit




'release objects
	Set objExcel = Nothing
	Set objWorkbook = Nothing