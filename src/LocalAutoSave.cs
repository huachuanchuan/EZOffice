using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace EZOffice
{
    internal static class LocalAutoSave
    {
        public static string SaveChangedFiles(dynamic documents, Func<bool> isEnabled)
        {
            if (documents == null)
                return null;

            int count = (int)documents.Count;
            if (count == 0)
                return "等待打开文件";

            // Save events can change the collection of open files.
            var openDocuments = new List<object>(count);
            for (int i = 1; i <= count; i++)
                openDocuments.Add((object)documents[i]);

            List<string> syncFolders = GetSyncFolders();
            int saved = 0;
            int failed = 0;
            bool needsFirstSave = false;
            bool skipped = false;
            foreach (dynamic document in openDocuments)
            {
                if (!isEnabled())
                    break;
                try
                {
                    if (Convert.ToBoolean((object)document.Saved))
                        continue;
                    if (string.IsNullOrEmpty((string)document.Path))
                    {
                        needsFirstSave = true;
                        continue;
                    }
                    if (Convert.ToBoolean((object)document.ReadOnly) || !IsLocalFile((string)document.FullName, syncFolders))
                    {
                        skipped = true;
                        continue;
                    }

                    document.Save();
                    if (Convert.ToBoolean((object)document.Saved))
                        saved++;
                    else
                        failed++;
                }
                catch (COMException error)
                {
                    Trace.WriteLine(error);
                    failed++;
                }
            }

            if (failed > 0)
                return string.Format("已保存 {0} 份，{1} 份待重试", saved, failed);
            if (needsFirstSave)
                return saved > 0 ? "已保存本地文件，新文件请先另存" : "新文件请先按 Ctrl+S 保存到本地";
            if (saved > 0)
                return string.Format("{0:HH:mm} 已保存 {1} 份", DateTime.Now, saved);
            return skipped ? "已跳过只读或非本地文档" : null;
        }

        private static bool IsLocalFile(string file, List<string> syncFolders)
        {
            // Only drive-letter paths on a local volume can reach Document.Save.
            if (file.Length < 4 || file[1] != ':' || file[2] != '\\')
                return false;
            uint driveType = GetDriveType(file.Substring(0, 3));
            if (driveType != 2 && driveType != 3 && driveType != 6)
                return false;

            foreach (string folder in syncFolders)
            {
                if (file.StartsWith(folder, StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            // A removed or renamed file would otherwise risk opening Save As.
            return File.Exists(file);
        }

        private static List<string> GetSyncFolders()
        {
            var folders = new List<string>();
            AddSyncFolder(folders, Environment.GetEnvironmentVariable("OneDrive"));
            AddSyncFolder(folders, Environment.GetEnvironmentVariable("OneDriveConsumer"));
            AddSyncFolder(folders, Environment.GetEnvironmentVariable("OneDriveCommercial"));
            try
            {
                using (RegistryKey accounts = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\OneDrive\Accounts"))
                {
                    if (accounts != null)
                    {
                        foreach (string name in accounts.GetSubKeyNames())
                        {
                            using (RegistryKey account = accounts.OpenSubKey(name))
                                AddSyncFolder(folders, account == null ? null : account.GetValue("UserFolder") as string);
                        }
                    }
                }
            }
            catch (SystemException error) { Trace.WriteLine(error); }
            return folders;
        }

        private static void AddSyncFolder(List<string> folders, string path)
        {
            if (!string.IsNullOrWhiteSpace(path))
                folders.Add(path.TrimEnd('\\', '/') + "\\");
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern uint GetDriveType(string rootPath);
    }
}
