using System;
using System.Diagnostics;
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
        Console.WriteLine("Enter folder's path which contains all files:");
        var linkFiles = Console.ReadLine();
        while (linkFiles == null || !Directory.Exists(linkFiles))
        {
            Console.WriteLine("Enter folder's path again:");
            linkFiles = Console.ReadLine();
        }

        Console.WriteLine("Enter the path of excel file:");
        var excelFile = Console.ReadLine();
        while (excelFile == null || !File.Exists(excelFile))
        {
            Console.WriteLine("Enter the excel file's path again:");
            linkFiles = Console.ReadLine();
        }

        // Get all files
        var allFiles = Directory.GetFiles(linkFiles).Select(Path.GetFileName).ToList();

        // Get all folders
        var allFolders = Directory.GetDirectories(linkFiles).Select(Path.GetFileName).ToList();

        if (allFiles == null || allFiles.Count == 0)
            return;

        if (allFolders != null && allFolders.Count > 0)
        {
            //Handle name files is same of name folder
            HandleSameName(linkFiles, allFiles, allFolders);
        }

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
    var listNotExistFile = new List<string>();
    try
    {
        // Kill excel file if opening
        KillExcelProcessByName(excelFile);

        // Process
        workbook = excelApp.Workbooks.Open(excelFile);
        Worksheet worksheet = (Worksheet)workbook.Sheets[1];
        int rowCount = worksheet.UsedRange.Rows.Count;
        for (int row = 2; row <= rowCount; row++)
        {
            string? folder = worksheet.Cells[row, 1].Value2.ToString();
            string? file = worksheet.Cells[row, 2].Value2 as string;

            if (folder == null || file == null)
                continue;

            // Check existion of file to show warning message
            var fileLink = Path.Combine(rootLink, file);
            if (!File.Exists(fileLink))
            {
                listNotExistFile.Add(file);
                continue;
            }

            // Create new folder if not exist
            var folderLink = Path.Combine(rootLink, folder);
            if (!Directory.Exists(folderLink))
            {
                Directory.CreateDirectory(folderLink);
            }

            var destFolderLink = Path.Combine(rootLink, folder, file);
            File.Move(fileLink, destFolderLink, true);
        }

        // Show list of not exist files
        if(listNotExistFile.Count > 0)
        {
            string message = "";
            listNotExistFile.ForEach(x => message += $"{x}\n");
            MessageBox.Show(message, "List of not exist files in excel file", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
    finally
    {
        var misValue = System.Reflection.Missing.Value;
        workbook?.Close(false, misValue, misValue);
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

void KillExcelProcessByName(string fileName)
{
    Process[] processes = Process.GetProcessesByName("EXCEL");
    foreach (Process process in processes)
    {
        try
        {
            // Get the file name of the Excel workbook opened by the process
            string openedFileName = process.MainWindowTitle;

            // Check if the file name matches the target file
            if (openedFileName.Equals(fileName, StringComparison.OrdinalIgnoreCase))
            {
                // Kill the Excel process
                process.Kill();
                return;
            }
        }
        catch (Exception)
        {
            // Ignore any exceptions when accessing the process
        }
    }
}

Main();