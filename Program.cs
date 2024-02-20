namespace excellabel;

using Microsoft.Office.Core;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Extensions.Configuration;

class Program
{
	static LabelInfo CreateNewLabel(SensitivityLabel sensitivity, IConfiguration config)
	{
		if (sensitivity == null) throw new System.ArgumentException("An invalid value for sensitivity was supplied");

		LabelInfo label = sensitivity.CreateLabelInfo();
		label.LabelName = config["settings:labelName"];
		label.LabelId = config["settings:labelId"];
		label.Justification = "No longer applies";
		label.AssignmentMethod = MsoAssignmentMethod.AUTO;
		label.SiteId = config["settings:siteId"];

		return label;
	}

	[MTAThread]
	static void Main(string[] args)
	{
		IConfiguration config = new ConfigurationBuilder()
			.AddJsonFile("appsettings.json")
			.AddJsonFile("appsettings.development.json", true)
			.Build();

		Console.WriteLine("Excel Label Test Program");

		Console.WriteLine("Launching Microsoft Excel");
		Excel.Application exApp = new Excel.Application();

		Console.WriteLine("Creating Workbook");
		Excel.Workbook wb = exApp.Workbooks.Add();

		exApp.Visible = true;

		if (exApp.Ready) Console.WriteLine("App Ready");

		Console.WriteLine("Sleeping for 5 seconds...");
		Thread.Sleep(5000);

		Console.WriteLine("Reading sensitivity label information");
		SensitivityLabel workbookLabel = wb.SensitivityLabel;

		Console.WriteLine(workbookLabel.SensitivityLabelError.ToString());

		workbookLabel.LabelChanged += (LabelInfo oldLabelInfo, LabelInfo newLabelInfo, int hResult, object Context) => {
			Console.WriteLine("Label Changed");
			Console.WriteLine($"LabelName        : {newLabelInfo.LabelName}        (was {oldLabelInfo.LabelName})");
			Console.WriteLine($"LabelId          : {newLabelInfo.LabelId}          (was {oldLabelInfo.LabelId})");
			Console.WriteLine($"SiteId           : {newLabelInfo.SiteId}           (was {oldLabelInfo.SiteId})");
			Console.WriteLine($"AssignmentMethod : {newLabelInfo.AssignmentMethod} (was {oldLabelInfo.AssignmentMethod})");
			Console.WriteLine($"Justification    : {newLabelInfo.Justification}    (was {oldLabelInfo.Justification})");
		};

		var li = workbookLabel.GetLabel();
		var label = CreateNewLabel(workbookLabel, config);

		workbookLabel.SetLabel(label, null);

		Console.WriteLine("Press enter to exit");
		Console.ReadLine();

		exApp.Quit();
	}
}