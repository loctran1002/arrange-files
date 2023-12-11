using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

#region General
string nameSheet = "Formated";
#endregion

List<string> listUsedFile = new List<string>();
List<string> listNotExistFile = new List<string>();

void Main()
{
    try
    {
        // Start processing
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
            excelFile = Console.ReadLine();
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

        // Remove the used files
        RemoveUsedFiles(listUsedFile);

        // Show list of not exist files
        if (listNotExistFile.Count > 0)
        {
            string message = "";
            listNotExistFile.ForEach(x => message += $"{x}\n");
            MessageBox.Show(message, "List of not exist files in excel file", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        // Reset data
        listUsedFile.Clear();
        listNotExistFile.Clear();
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
        KillExcelProcessByName(Path.GetFileName(excelFile));

        // Process
        workbook = excelApp.Workbooks.Open(excelFile);
        Worksheet worksheet = (Worksheet)workbook.Sheets[nameSheet];
        int rowCount = worksheet.UsedRange.Rows.Count;
        for (int row = 2; row <= rowCount; row++)
        {
            string? folder = worksheet.Cells[row, 1].Value2 as string;
            string? file = worksheet.Cells[row, 2].Value2 as string;

            if (folder == null)
                continue;
            else folder = folder.TrimEnd();

            // Create new folder if not exist
            var folderLink = Path.Combine(rootLink, folder);
            if (!Directory.Exists(folderLink))
            {
                Directory.CreateDirectory(folderLink);
            }

            if (file == null || file == "" || file == "\n")
                continue;
            else file = file.TrimEnd();

            // Check existion of file to show warning message
            List<string?> listPath = CreateFilePath(rootLink, allFiles, file);
            foreach (var fileLink in listPath)
            {
                if (!File.Exists(fileLink))
                {
                    listNotExistFile.Add(file);
                    continue;
                }

                var destFolderLink = Path.Combine(rootLink, folder, Path.GetFileName(fileLink));
                File.Copy(fileLink, destFolderLink, true);
                listUsedFile.Add(fileLink);
            }
        }
    }
    finally
    {
        var misValue = System.Reflection.Missing.Value;
        workbook?.Close(false, misValue, misValue);
        excelApp.Quit();
        //KillProcessById(excelApp.Hwnd);
    }
}

List<string?> CreateFilePath(string rootLink, List<string?> allFiles, string file)
{
    var listPath = new List<string?>();
    foreach (var nameRealFile in allFiles)
    {
        if (nameRealFile == null)
            continue;

        if (nameRealFile.Contains(file))
        {
            listPath.Add(Path.Combine(rootLink, nameRealFile));
        }

        //// Check first characters
        //var formatedName = nameRealFile;
        //var firstChars = formatedName.Split('_').First();
        //if (firstChars == null)
        //    continue; // TODO
        //if (IsNumeric(firstChars))
        //{
        //    formatedName = formatedName.Remove(0, formatedName.IndexOf('_') + 1);
        //}
        //else if(file == formatedName)
        //{
        //    name = nameRealFile;
        //    break;
        //}

        //// Remove last characters
        //var lastIndex = nameRealFile.LastIndexOf('_');
        //if (lastIndex == -1)
        //    continue;
        //formatedName = formatedName.Remove(lastIndex);

        //// Find real file's name
        //if (file == formatedName)
        //{
        //    name = nameRealFile;
        //    break;
        //}
    }

    return listPath;
}

bool IsNumeric(string input)
{
    return Regex.IsMatch(input, @"^\d+$");
}

void RemoveUsedFiles(List<string> listFile)
{
    foreach (var fileLink in listFile)
    {
        if (File.Exists(fileLink))
            File.Delete(fileLink);
    }
}

void HandleSameName(string rootLink, List<string> allFiles, List<string> allFolders)
{
    foreach (var file in allFiles)
    {
        string nameFile = Path.GetFileNameWithoutExtension(file);
        var indexFolder = allFolders.FindIndex(x => x == nameFile);
        if (indexFolder == -1)
            continue;

        var fileLink = Path.Combine(rootLink, file);
        var destFolderLink = Path.Combine(rootLink, allFolders[indexFolder], file);
        File.Move(fileLink, destFolderLink, true);
        //File.Copy(fileLink, destFolderLink, true);

        // Add file to remove after end program
        //listUsedFile.Add(fileLink);
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
            if (openedFileName == $"{fileName} - Excel")
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

void KillProcessById(int processId)
{
    try
    {
        // Get the process by ID
        Process excelProcess = Process.GetProcessById(processId);

        // Kill the process
        excelProcess.Kill();

        // Release resources
        excelProcess.Close();
    }
    catch (ArgumentException)
    {
        // Process with the specified ID is not running
        Console.WriteLine($"Process with ID {processId} is not currently running.");
    }
}

Main();