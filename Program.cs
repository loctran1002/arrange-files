using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

void Main()
{
	try
	{
		//while (true)
		//{
		Console.WriteLine("Enter link to folder which contains all files:");
		var linkFiles = Console.ReadLine();
		while (linkFiles == null || !Directory.Exists(linkFiles))
		{
			Console.WriteLine("Enter again:");
			linkFiles = Console.ReadLine();
		}

		Console.WriteLine("Enter excel file:");
		var excelFile = Console.ReadLine();
		while (linkFiles == null || !Directory.Exists(linkFiles))
		{
			Console.WriteLine("Enter again:");
			linkFiles = Console.ReadLine();
		}

		// Get all files
		var allFiles = Directory.GetFiles(linkFiles).Select(Path.GetFileName).ToList();

		// Get all folders
		var allFolders = Directory.GetDirectories(linkFiles).Select(Path.GetFileName).ToList();

		if (allFiles == null || allFiles.Count == 0 || allFolders == null || allFolders.Count == 0)
			return;

		//Handle name files is same of name folder
		HandleSameName(linkFiles, allFiles, allFolders);

		//Handle arrange excel files
		HandleFileExcel(linkFiles, excelFile, allFiles, allFolders);

		Console.WriteLine("===================== Done =====================");
		//}
	}
	catch (Exception ex)
	{
		MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
	}
}

void HandleFileExcel(string rootLink, string excelFile, List<string?> allFiles, List<string?> allFolders)
{
	Application excelApp = new Application();
	Workbook workbook = null;
	try
	{
		workbook = excelApp.Workbooks.Open(excelFile);
		Worksheet worksheet = (Worksheet)workbook.Sheets[1];
		int rowCount = worksheet.UsedRange.Rows.Count;

		for (int row = 2; row <= rowCount; row++)
		{
			string? folder = worksheet.Cells[row, 1].Value2 as string;
			string? file = worksheet.Cells[row, 2].Value2 as string;

			if (folder == null || file == null)
				continue;

			var fileLink = Path.Combine(rootLink, file);
			var destFolderLink = Path.Combine(rootLink, folder, file);
			File.Move(fileLink, destFolderLink, true);
		}
	}
	finally
	{
		// Đóng workbook và ứng dụng Excel
		workbook?.Close();
		excelApp.Quit();
	}
}

void HandleSameName(string rootLink, List<string> allFiles, List<string> allFolders)
{
	foreach (var file in allFiles)
	{
		string nameFile = file.Split('.').First();
		var indexFolder = allFolders.FindIndex(x => x.Split('.').First() == nameFile);
		if (indexFolder == -1)
			continue;

		var fileLink = Path.Combine(rootLink, file);
		var destFolderLink = Path.Combine(rootLink, allFolders[indexFolder], file);
		File.Move(fileLink, destFolderLink, true);
	}
}

Main();