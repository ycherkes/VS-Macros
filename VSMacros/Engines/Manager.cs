//-----------------------------------------------------------------------
// <copyright file="Manager.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Xml.Linq;
using VSMacros.Dialogs;
using VSMacros.Interfaces;
using VSMacros.Models;
using Process = System.Diagnostics.Process;

namespace VSMacros.Engines
{
    public sealed class Manager
    {
        private static Manager instance;
        internal Executor executor;
        private Executor Executor => (executor ?? (executor = new Executor()));

        #region Paths

        private const string CurrentMacroFileName = "Current.js";
        private const string ShortcutsFileName = "Shortcuts.xml";
        public const string IntellisenseFileName = "dte.js";
        private const string FolderExpansionFileName = "FolderExpansion.xml";

        public static string MacrosPath => Path.Combine(VSMacrosPackage.Current.MacroDirectory, "Macros");

        public static string CurrentMacroPath => Path.Combine(MacrosPath, CurrentMacroFileName);

        public static string SamplesFolderPath => Path.Combine(MacrosPath, "Samples");

        public static string IntellisensePath => Path.Combine(VSMacrosPackage.Current.MacroDirectory, IntellisenseFileName);

        public static string ShortcutsPath => Path.Combine(VSMacrosPackage.Current.MacroDirectory, ShortcutsFileName);

        #endregion


        public static string[] Shortcuts { get; private set; }
        private bool shortcutsLoaded;
        private bool shortcutsDirty;

        private IServiceProvider serviceProvider;
        private IVsUIShell uiShell;
        private DTE dte;

        private IRecorder recorder;

        private MacroFSNode SelectedMacro => MacrosControl.Current != null ? MacrosControl.Current.SelectedNode : null;

        private Manager(IServiceProvider provider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            serviceProvider = provider;
            uiShell = (IVsUIShell)provider.GetService(typeof(SVsUIShell));
            dte = (DTE)provider.GetService(typeof(SDTE));
            recorder = (IRecorder)serviceProvider.GetService(typeof(IRecorder));

            LoadShortcuts();
            shortcutsLoaded = true;
            shortcutsDirty = false;

            CreateFileSystem();
        }

        private static void AttachEvents(Executor executor)
        {
            executor.ResetMessages();

            executor.Complete += (sender, eventInfo) =>
            {
                if (eventInfo.IsError)
                {
                    Instance.ShowMessageBox(eventInfo.ErrorMessage);
                }

                instance.Executor.IsEngineRunning = false;

                _ = ResetToolbar();
            };
        }

        private static async Task ResetToolbar()
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                VSMacrosPackage.Current.ClearStatusBar();
                VSMacrosPackage.Current.UpdateButtonsForPlayback(false);
            }
            catch (Exception)
            {
                // Visual Studio is closing during execution.
            }
        }

        public static Manager Instance => instance ?? (instance = new Manager(VSMacrosPackage.Current));

        public IVsWindowFrame PreviousWindow { get; set; }
        public bool PreviousWindowIsDocument { get; set; }

        public bool IsRecording { get; private set; }

        public void StartRecording()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Move focus back to previous window
            if (dte.ActiveWindow.Caption == "Macro Explorer" && PreviousWindow != null)
            {
                PreviousWindow.Show();
            }

            PreviousWindowIsDocument = dte.ActiveWindow.Kind == "Document";

            IsRecording = true;
            recorder.StartRecording();
        }


        public void StopRecording()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string current = CurrentMacroPath;

            bool currentWasOpen = false;

            // Close current macro if open
            try
            {
                dte.Documents.Item(CurrentMacroPath).Close(vsSaveChanges.vsSaveChangesNo);
                currentWasOpen = true;
            }
            catch (Exception e)
            {
                if (ErrorHandler.IsCriticalException(e))
                {
                    throw;
                }
            }

            recorder.StopRecording(current);
            IsRecording = false;

            MacroFSNode.SelectNode(CurrentMacroPath);

            // Reopen current macro
            if (currentWasOpen)
            {
                VsShellUtilities.OpenDocument(VSMacrosPackage.Current, CurrentMacroPath);
                PreviousWindow.Show();
            }
        }

        public void Playback(string path, int iterations = 1)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (string.IsNullOrEmpty(path))
            {
                path = SelectedMacro == null ? CurrentMacroPath : SelectedMacro.FullPath;
            }

            // Before playing back, save the macro file
            SaveMacroIfDirty(path);

            // Move focus to first window
            if (dte.ActiveWindow.Caption == "Macro Explorer" && PreviousWindow != null)
            {
                PreviousWindow.Show();
            }

            VSMacrosPackage.Current.StatusBarChange(Resources.StatusBarPlayingText, 1);

            TogglePlayback(path, iterations);
        }

        public void PlaybackMultipleTimes(string path)
        {
            PlaybackMultipleTimesDialog dlg = new PlaybackMultipleTimesDialog();
            bool? result = dlg.ShowDialog();

            if (!result.HasValue || !result.Value) return;
            if (int.TryParse(dlg.IterationsTextbox.Text, out _))
            {
                Playback(string.Empty, dlg.Iterations);
            }
        }

        private void TogglePlayback(string path, int iterations)
        {
            AttachEvents(Executor);

            if (Instance.executor.IsEngineRunning)
            {
                VSMacrosPackage.Current.ClearStatusBar();
                Instance.executor.StopEngine();
            }
            else
            {
                VSMacrosPackage.Current.UpdateButtonsForPlayback(true);
                Executor.RunEngine(iterations, path);
                instance.Executor.CurrentlyExecutingMacro = GetExecutingMacroNameForPossibleErrorDisplay(SelectedMacro, path);
            }
        }

        private static string GetExecutingMacroNameForPossibleErrorDisplay(MacroFSNode node, string path)
        {
            if (node != null)
            {
                return node.Name;
            }

            if (path == null)
            {
                throw new ArgumentNullException("path");
            }

            int lastBackslash = path.LastIndexOf('\\');
            string fileName = Path.GetFileNameWithoutExtension(path.Substring(lastBackslash != -1 ? lastBackslash + 1 : 0));

            return fileName;
        }

        public void PlaybackCommand(int cmd)
        {
            // Load shortcuts if not already loaded
            if (!shortcutsLoaded)
            {
                LoadShortcuts();
            }

            // Get path to macro bound to the shortcut
            string path = Shortcuts[cmd];

            if (!string.IsNullOrEmpty(path))
            {
                Playback(path);
            }
        }

        public void StopPlayback()
        {
        }

        public void OpenFolder(string path = null)
        {
            path = !string.IsNullOrEmpty(path) ? path : SelectedMacro.FullPath;

            // Open the macro directory and let the user manage the macros
            _ = Task.Run(() => Process.Start(path));
        }

        public void SaveCurrent()
        {
            SaveCurrentDialog dlg = new SaveCurrentDialog();
            dlg.ShowDialog();

            if (dlg.DialogResult == true)
            {
                try
                {
                    string pathToNew = Path.Combine(MacrosPath, dlg.MacroName.Text + ".js");
                    string pathToCurrent = CurrentMacroPath;

                    int newShortcutNumber = dlg.SelectedShortcutNumber;

                    // Move Current to new file and create a new Current
                    File.Move(pathToCurrent, pathToNew);
                    CreateCurrentMacro();

                    MacroFSNode macro = new MacroFSNode(pathToNew, MacroFSNode.RootNode);

                    if (newShortcutNumber != MacroFSNode.None)
                    {
                        // Update dictionary
                        Shortcuts[newShortcutNumber] = macro.FullPath;
                    }

                    SaveShortcuts(true);

                    Refresh();

                    // Select new node
                    MacroFSNode.SelectNode(pathToNew);
                }
                catch (Exception e)
                {
                    if (ErrorHandler.IsCriticalException(e))
                    {
                        throw;
                    }

                    ShowMessageBox(e.Message);
                }
            }
        }

        public void Refresh(bool reloadShortcut = true)
        {
            // If the shortcuts have been modified, ask to save them
            if (shortcutsDirty && reloadShortcut)
            {
                VSConstants.MessageBoxResult result = ShowMessageBox(Resources.ShortcutsChanged, OLEMSGBUTTON.OLEMSGBUTTON_YESNOCANCEL);

                switch (result)
                {
                    case VSConstants.MessageBoxResult.IDCANCEL:
                        return;
                    case VSConstants.MessageBoxResult.IDYES:
                        SaveShortcuts();
                        break;
                }
            }

            // Recreate file system to ensure that the required files exist
            CreateFileSystem();

            MacroFSNode.RefreshTree();

            LoadShortcuts();
        }

        public void Edit()
        {
            // TODO detect when a macro is dragged and it's opened -> use the overload to get the itemID
            MacroFSNode macro = SelectedMacro;
            string path = macro.FullPath;

            VsShellUtilities.OpenDocument(VSMacrosPackage.Current, path);
        }

        public void Rename()
        {
            MacroFSNode macro = SelectedMacro;

            if (macro.FullPath != CurrentMacroPath)
            {
                macro.EnableEdit();
            }
        }

        public void AssignShortcut()
        {
            AssignShortcutDialog dlg = new AssignShortcutDialog();
            dlg.ShowDialog();

            if (dlg.DialogResult == true)
            {
                MacroFSNode macro = SelectedMacro;

                // Remove old shortcut if it exists
                if (macro.Shortcut != MacroFSNode.None)
                {
                    Shortcuts[macro.Shortcut] = string.Empty;
                }

                int newShortcutNumber = dlg.SelectedShortcutNumber;

                // At this point, the shortcut has been removed
                // Assign a new one only if the user selected a key binding
                if (newShortcutNumber != MacroFSNode.None)
                {
                    // Get the node that previously owned that shortcut
                    MacroFSNode previousNode = MacroFSNode.FindNodeFromFullPath(Shortcuts[newShortcutNumber]);

                    // Update dictionary
                    Shortcuts[newShortcutNumber] = macro.FullPath;

                    // Update the UI binding for the old node
                    previousNode.Shortcut = MacroFSNode.ToFetch;
                }

                // Update UI with new macro's shortcut
                macro.Shortcut = MacroFSNode.ToFetch;

                // Mark the shortcuts in memory as dirty
                shortcutsDirty = true;
            }
        }

        public void SetShortcutsDirty()
        {
            shortcutsDirty = true;
        }

        public void Delete()
        {
            MacroFSNode macro = SelectedMacro;

            // Don't delete if macro is being edited
            if (macro.IsEditable)
            {
                return;
            }

            string path = macro.FullPath;

            FileSystemInfo file;
            string fileName = Path.GetFileNameWithoutExtension(path);
            string message;

            if (macro.IsDirectory)
            {
                file = new DirectoryInfo(path);
                message = string.Format(Resources.DeleteFolder, fileName);
            }
            else
            {
                file = new FileInfo(path);
                message = string.Format(Resources.DeleteMacro, fileName);
            }

            if (file.Exists)
            {
                VSConstants.MessageBoxResult result;
                result = ShowMessageBox(message, OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL);

                if (result == VSConstants.MessageBoxResult.IDOK)
                {
                    try
                    {
                        // Delete file or directory from disk
                        DeleteFileOrFolder(path);

                        // Delete file from collection
                        macro.Delete();
                    }
                    catch (Exception e)
                    {
                        ShowMessageBox(e.Message);
                    }
                }
            }
            else
            {
                macro.Delete();
            }
        }

        public void NewMacro()
        {
            MacroFSNode macro = SelectedMacro;
            macro.IsExpanded = true;

            string basePath = Path.Combine(macro.FullPath, "New Macro");
            string extension = ".js";

            string path = basePath + extension;

            // Increase count until filename is available (e.g. 'New Macro (2).js')
            int count = 2;
            while (File.Exists(path))
            {
                path = basePath + " (" + count++ + ")" + extension;
            }

            // Create the file
            File.WriteAllText(path, "/// <reference path=\"" + IntellisensePath + "\" />");

            // Refresh the tree
            MacroFSNode.RefreshTree();

            // Select new node
            MacroFSNode node = MacroFSNode.SelectNode(path);
            node.IsExpanded = true;
            node.IsEditable = true;
        }

        public void NewFolder()
        {
            MacroFSNode macro = SelectedMacro;
            macro.IsExpanded = true;

            string basePath = Path.Combine(macro.FullPath, "New Folder");
            string path = basePath;

            int count = 2;
            while (Directory.Exists(path))
            {
                path = basePath + " (" + count++ + ")";
            }

            Directory.CreateDirectory(path);
            MacroFSNode.RefreshTree();

            MacroFSNode node = MacroFSNode.SelectNode(path);
            node.IsEditable = true;
        }

        public static void CreateFileSystem()
        {
            // Create main macro directory
            if (!Directory.Exists(VSMacrosPackage.Current.MacroDirectory))
            {
                Directory.CreateDirectory(VSMacrosPackage.Current.MacroDirectory);
            }

            // Create macros folder directory
            if (!Directory.Exists(MacrosPath))
            {
                Directory.CreateDirectory(MacrosPath);
            }

            // Create current macro file
            CreateCurrentMacro();

            // Create shortcuts file
            CreateShortcutFile();

            // Copy Samples folder
            string samplesTargetDir = SamplesFolderPath;
            if (!Directory.Exists(samplesTargetDir))
            {
                string samplesSourceDir = Path.Combine(VSMacrosPackage.Current.AssemblyDirectory, "Macros", "Samples");
                DirectoryCopy(samplesSourceDir, samplesTargetDir, true);
            }

            // Copy DTE IntelliSense file
            string dteFileTargetPath = IntellisensePath;
            if (!File.Exists(dteFileTargetPath))
            {
                string dteFileSourcePath = Path.Combine(VSMacrosPackage.Current.AssemblyDirectory, "Intellisense", IntellisenseFileName);
                File.Copy(dteFileSourcePath, dteFileTargetPath);
            }
        }

        public void Close()
        {
            SaveFolderExpansion();
            SaveShortcuts();
        }

        private string RelativeIntellisensePath(int depth)
        {
            string path = IntellisenseFileName;

            for (int i = 0; i < depth; i++)
            {
                path = "../" + path;
            }

            return path;
        }

        public void MoveItem(MacroFSNode sourceItem, MacroFSNode targetItem)
        {
            string sourcePath = sourceItem.FullPath;
            string targetPath = Path.Combine(targetItem.FullPath, sourceItem.Name);
            string extension = ".js";

            MacroFSNode selected;

            // We want to expand the node and all its parents if it was expanded before OR if it is a file
            bool wasExpanded = sourceItem.IsExpanded;

            try
            {
                // Move on disk
                if (sourceItem.IsDirectory)
                {
                    Directory.Move(sourcePath, targetPath);
                }
                else
                {
                    targetPath = targetPath + extension;
                    File.Move(sourcePath, targetPath);

                    // Close in the editor
                    Reopen(sourcePath, targetPath);
                }

                // Move shortcut as well
                if (sourceItem.Shortcut != MacroFSNode.None)
                {
                    int shortcutNumber = sourceItem.Shortcut;
                    Shortcuts[shortcutNumber] = targetPath;
                }
            }
            catch (Exception e)
            {
                if (ErrorHandler.IsCriticalException(e))
                {
                    throw;
                }

                targetPath = sourceItem.FullPath;

                Instance.ShowMessageBox(e.Message);
            }

            CreateCurrentMacro();

            // Refresh tree
            MacroFSNode.RefreshTree();

            // Restore previously selected node
            selected = MacroFSNode.SelectNode(targetPath);
            selected.IsExpanded = wasExpanded;
            selected.Parent.IsExpanded = true;

            // Notify change in shortcut
            selected.Shortcut = MacroFSNode.ToFetch;

            // Make editable if the macro is the current macro
            if (sourceItem.FullPath == CurrentMacroPath)
            {
                selected.IsEditable = true;
            }
        }

        #region Helper Methods

        private void Reopen(string source, string target)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                dte.Documents.Item(source).Close(vsSaveChanges.vsSaveChangesNo);
                dte.ItemOperations.OpenFile(target);
            }
            catch (ArgumentException) { }
        }

        private StreamReader LoadFile(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    throw new FileNotFoundException(Resources.MacroNotFound);
                }

                StreamReader str = new StreamReader(path);

                return str;
            }
            catch (FileNotFoundException e)
            {
                ShowMessageBox(e.Message);
            }

            return null;
        }

        private void SaveMacro(Stream str, string path)
        {
            try
            {
                using (var fileStream = File.Create(path))
                {
                    str.Seek(0, SeekOrigin.Begin);
                    str.CopyTo(fileStream);
                }
            }
            catch (Exception e)
            {
                ShowMessageBox(e.Message);
            }
        }

        public void LoadFolderExpansion()
        {
            string path = Path.Combine(VSMacrosPackage.Current.UserLocalDataPath, FolderExpansionFileName);

            if (File.Exists(path))
            {
                HashSet<string> enabledDirs = new HashSet<string>();

                string[] folders = File.ReadAllText(path).Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var s in folders)
                {
                    enabledDirs.Add(s);
                }

                MacroFSNode.EnabledDirectories = enabledDirs;
            }
        }

        private void SaveFolderExpansion()
        {
            var folders = string.Join(Environment.NewLine, MacroFSNode.EnabledDirectories);

            File.WriteAllText(Path.Combine(VSMacrosPackage.Current.UserLocalDataPath, FolderExpansionFileName), folders);
        }

        private void LoadShortcuts()
        {
            shortcutsLoaded = true;

            try
            {
                // Get the path to the shortcut file
                string path = ShortcutsPath;

                // If the file doesn't exist, initialize the Shortcuts array with empty strings
                if (!File.Exists(path))
                {
                    Shortcuts = Enumerable.Repeat(string.Empty, 10).ToArray();
                }
                else
                {
                    // Otherwise, load it
                    // Load XML file
                    var root = XDocument.Load(path);

                    // Parse to dictionary
                    Shortcuts = root.Descendants("command")
                                          .Select(elmt => elmt.Value)
                                          .ToArray();
                }
            }
            catch (Exception e)
            {
                ShowMessageBox(e.Message);
            }
        }

        public void SaveShortcuts(bool overwrite = false)
        {
            if (shortcutsDirty || overwrite)
            {
                XDocument xmlShortcuts =
                    new XDocument(
                        new XDeclaration("1.0", "utf-8", "yes"),
                        new XElement("commands",
                            from s in Shortcuts
                            select new XElement("command",
                                new XText(s))));

                xmlShortcuts.Save(ShortcutsPath);

                shortcutsDirty = false;
            }
        }

        private static void CreateCurrentMacro()
        {
            if (!File.Exists(CurrentMacroPath))
            {
                File.Create(CurrentMacroPath).Close();

                Task.Run(() =>
                {
                    // Write to current macro file
                    File.WriteAllText(CurrentMacroPath, "/// <reference path=\"" + IntellisensePath + "\" />");
                });
            }
        }

        private static void CreateShortcutFile()
        {
            string shortcutsPath = ShortcutsPath;
            if (!File.Exists(shortcutsPath))
            {
                // Create file for writing UTF-8 encoded text
                File.WriteAllText(shortcutsPath, "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><commands><command>Command not bound. Do not use.</command><command/><command/><command/><command/><command/><command/><command/><command/><command/></commands>");
            }
        }

        private void SaveMacroIfDirty(string path)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                Document doc = dte.Documents.Item(path);

                if (!doc.Saved)
                {
                    if (VSConstants.MessageBoxResult.IDYES == ShowMessageBox(
                        string.Format(Resources.MacroNotSavedBeforePlayback, Path.GetFileNameWithoutExtension(path)),
                        OLEMSGBUTTON.OLEMSGBUTTON_YESNO))
                    {
                        doc.Save();
                    }
                }
            }
            catch (Exception e)
            {
                if (ErrorHandler.IsCriticalException(e)) { throw; }
            }
        }

        public VSConstants.MessageBoxResult ShowMessageBox(string message, OLEMSGBUTTON btn = OLEMSGBUTTON.OLEMSGBUTTON_OK)
        {
            if (uiShell == null)
            {
                return VSConstants.MessageBoxResult.IDABORT;
            }

            Guid clsid = Guid.Empty;
            int result;

            ErrorHandler.ThrowOnFailure(
              uiShell.ShowMessageBox(
                0,
                ref clsid,
                string.Empty,
                message,
                string.Empty,
                0,
                btn,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                OLEMSGICON.OLEMSGICON_WARNING,
                0,        // false
                out result));

            return (VSConstants.MessageBoxResult)result;
        }

        #endregion

        #region Disk Operations

        private const int FO_DELETE = 0x0003;
        private const int FOF_ALLOWUNDO = 0x0040;           // Preserve undo information, if possible. 
        private const int FOF_NOCONFIRMATION = 0x0010;      // Show no confirmation dialog box to the user

        // Struct which contains information that the SHFileOperation function uses to perform file operations. 
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SHFILEOPSTRUCT
        {
            public IntPtr hwnd;
            [MarshalAs(UnmanagedType.U4)]
            public int wFunc;
            public string pFrom;
            public string pTo;
            public short fFlags;
            [MarshalAs(UnmanagedType.Bool)]
            public bool fAnyOperationsAborted;
            public IntPtr hNameMappings;
            public string lpszProgressTitle;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        static extern int SHFileOperation(ref SHFILEOPSTRUCT FileOp);

        public static void DeleteFileOrFolder(string path)
        {
            SHFILEOPSTRUCT fileop = new SHFILEOPSTRUCT();
            fileop.wFunc = FO_DELETE;
            fileop.pFrom = path + '\0' + '\0';
            fileop.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION;
            SHFileOperation(ref fileop);
        }

        public static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory doesn't exist, create it. 
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);

                if (!File.Exists(temppath))
                {
                    file.CopyTo(temppath, false);
                }
            }

            // If copying subdirectories, copy them and their contents to new location. 
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

        #endregion
    }
}
