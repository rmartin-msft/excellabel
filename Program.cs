namespace excellabel;

using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
	static void Main(string[] args)
	{
		Console.WriteLine("Hello, World!");

		Excel.Application exApp = new Excel.Application();		
		Excel.Workbook wb = exApp.Workbooks.Add();		

		var x = wb.SensitivityLabel;

		var li = x.GetLabel();

		Console.WriteLine(li.LabelName);
		Console.WriteLine(li.LabelId);
		
		exApp.Visible = true;

		var label = x.CreateLabelInfo();
		label.LabelName = "Public";
		label.LabelId = "87867195-f2b8-4ac2-b0b6-6bb73cb33afc";
		label.Justification = "No longer applies";
		label.AssignmentMethod = MsoAssignmentMethod.AUTO;

		x.SetLabel(label, null);		
	}
}
