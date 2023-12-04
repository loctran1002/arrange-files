using System;
using System.IO;
using System.Windows.Forms;

void Main()
{
    try
    {
        while (true)
        {
            Console.WriteLine("Enter link to folder which contains all files:");
            var linkFiles = Console.ReadLine();
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
                continue;

            // Handle name files is same of name folder
            //HandleSameName(linkFiles, allFiles, allFolders);

            var fileLink = @"D:\files\A_BC_12.abc.docx";
            var destFolderLink = @"D:\files\A_BC_12.xyz";
            File.Move(fileLink, destFolderLink, true);

            //foreach (var nameFile in allFolders)
            //{
            //    Console.WriteLine(Path.GetFileName(nameFile));
            //}

            Console.WriteLine("==============================================");
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        var destFolderLink = Path.Combine(rootLink, allFolders[indexFolder]);
        File.Move(fileLink, destFolderLink, true);

        //var fileBytes = File.ReadAllBytes(fileLink);
        //File.WriteAllBytes(destFolderLink, fileBytes);
        //File.Delete(fileLink);
    }
}

Main();