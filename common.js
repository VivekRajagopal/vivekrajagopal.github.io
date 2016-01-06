function openExcel()
	{
	try
		{
		objExcel = new ActiveXObject("Excel.Application");
		objExcel.Visible = true;
		objExcel.Workbooks.Add();
		//objExcel.SaveAs("hello.xlsx");
		objExcel.Close();
		}
	catch(error)
		{
		//alert(error.message);
		WScript.echo(error.message);
		}
	}
openExcel();