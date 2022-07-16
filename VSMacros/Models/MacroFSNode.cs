//-----------------------------------------------------------------------
// <copyright file="MacroFSNode.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media.Imaging;
using VSMacros.Engines;
using GelUtilities = Microsoft.Internal.VisualStudio.PlatformUI.Utilities;

namespace VSMacros.Models
{
    public sealed class MacroFSNode : INotifyPropertyChanged
    {
        // HashSet containing the enabled directories
        private static HashSet<string> enabledDirectories = new HashSet<string>();
        public static HashSet<string> EnabledDirectories
        {
            get => enabledDirectories;

            set
            {
                enabledDirectories = value;
                RootNode.SetIsExpanded(RootNode, enabledDirectories);
            }
        }

        // Properties the binding client watches
        private string fullPath;
        private int shortcut;
        private bool isEditable;
        private bool isExpanded;
        private bool isSelected;
        private bool isMatch;

        private readonly MacroFSNode parent;
        private ObservableCollection<MacroFSNode> children;

        // Constants
        public const int ToFetch = -1;
        public const int None = 0;
        public const string ShortcutKeys = "(CTRL+M, {0})";

        // Static members
        public static MacroFSNode RootNode { get; set; }
        private static bool Searching = false;

        // For INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        public MacroFSNode(string path, MacroFSNode parent = null)
        {
            IsDirectory = (File.GetAttributes(path) & FileAttributes.Directory) == FileAttributes.Directory;
            FullPath = path;
            shortcut = ToFetch;
            isEditable = false;
            isSelected = false;
            isExpanded = false;
            isMatch = false;
            this.parent = parent;

            // Monitor that node 
            //FileChangeMonitor.Instance.MonitorFileSystemEntry(this.FullPath, this.IsDirectory);
        }

        public string FullPath
        {
            get => fullPath;

            private set
            {
                fullPath = value;
                NotifyPropertyChanged("FullPath");
                NotifyPropertyChanged("Name");
            }
        }

        public string Name
        {
            get
            {
                string path = Path.GetFileNameWithoutExtension(FullPath);

                return string.IsNullOrWhiteSpace(path) ? FullPath : path;
            }

            set
            {
                try
                {
                    // Path.GetFullPath will throw an exception if the path is invalid
                    Path.GetFileName(value);

                    if (value != Name && !string.IsNullOrWhiteSpace(value))
                    {
                        string oldFullPath = FullPath;
                        string newFullPath = Path.Combine(Path.GetDirectoryName(FullPath), value + Path.GetExtension(FullPath));

                        // Update file system
                        if (IsDirectory)
                        {
                            Directory.Move(oldFullPath, newFullPath);

                            if (enabledDirectories.Remove(oldFullPath))
                            {
                                enabledDirectories.Add(newFullPath);
                            }
                        }
                        else
                        {
                            File.Move(oldFullPath, newFullPath);
                        }

                        // Update object
                        FullPath = newFullPath;

                        // Update shortcut
                        if (Shortcut >= None)
                        {
                            Manager.Shortcuts[shortcut] = newFullPath;
                            Manager.Instance.SaveShortcuts(true);
                        }
                    }
                }
                catch (Exception e)
                {
                    if (e.Message != null)
                    {
                        Manager.Instance.ShowMessageBox(e.Message);
                        RefreshTree();
                    }
                }
            }
        }

        public int Shortcut
        {
            get => shortcut;

            set
            {
                // Shortcut will be refetched
                shortcut = ToFetch;

                // Just notify the binding
                NotifyPropertyChanged("Shortcut");
                NotifyPropertyChanged("FormattedShortcut");
            }
        }

        public string FormattedShortcut
        {
            get
            {
                if (shortcut == ToFetch)
                {

                    shortcut = None;

                    // Find shortcut, if it exists
                    for (int i = 1; i < 10; i++)
                    {
                        if (string.Compare(Manager.Shortcuts[i], FullPath, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            shortcut = i;
                        }
                    }
                }

                if (shortcut != None)
                {
                    return string.Format(ShortcutKeys, shortcut);
                }

                return string.Empty;
            }
        }

        public BitmapSource Icon
        {
            get
            {
                if (IsDirectory)
                {
                    Bitmap bmp;

                    if (this == RootNode)
                    {
                        bmp = Resources.RootIcon;
                    }
                    else if (isExpanded)
                    {
                        bmp = Resources.FolderOpenedIcon;
                    }
                    else
                    {
                        bmp = Resources.FolderClosedIcon;
                    }

                    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                            bmp.GetHbitmap(),
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                }
                ThreadHelper.ThrowIfNotOnUIThread();
                IVsImageService2 imageService = (IVsImageService2)((IServiceProvider)VSMacrosPackage.Current).GetService(typeof(SVsImageService));
                if (imageService != null)
                {
                    //IVsUIObject uiObject = imageService.GetIconForFile(Path.GetFileName(this.FullPath), __VSUIDATAFORMAT.VSDF_WPF);
                    ImageMoniker imageMoniker = imageService.GetImageMonikerForFile(Path.GetFileName(FullPath));
                    IVsUIObject uiObject = imageService.GetImage(imageMoniker, new ImageAttributes
                    {
                        Format = (uint)__VSUIDATAFORMAT.VSDF_WPF,
                        Flags = (uint)_ImageAttributesFlags.IAF_RequiredFlags,
                        StructSize = Marshal.SizeOf(typeof(ImageAttributes)),
                        ImageType = (uint)_UIImageType.IT_Bitmap,
                        LogicalHeight = 50,
                        LogicalWidth = 50
                    });

                    if (uiObject != null)
                    {
                        BitmapSource bitmapSource = GelUtilities.GetObjectData(uiObject) as BitmapSource;
                        return bitmapSource;
                    }
                }

                // Would it be better to have a sane default image?
                return null;
            }
        }

        public bool IsEditable
        {
            get => isEditable;

            set
            {
                isEditable = value;
                NotifyPropertyChanged("IsEditable");
            }
        }

        public bool IsSelected
        {
            get => isSelected;

            set
            {
                isSelected = value;
                NotifyPropertyChanged("IsSelected");
            }
        }

        public bool IsExpanded
        {
            get => isExpanded || isMatch;

            set
            {
                if (!IsDirectory) return;

                isExpanded = value;

                if (IsExpanded)
                {

                    enabledDirectories.Add(FullPath);

                    // Expand parent as well
                    if (parent != null)
                    {
                        parent.IsExpanded = true;
                    }
                }
                else
                {

                    enabledDirectories.Remove(FullPath);
                }

                NotifyPropertyChanged("IsExpanded");
                NotifyPropertyChanged("Icon");
            }
        }

        public bool IsMatch
        {
            get => // If searching is not enabled, always return true
                   !Searching || isMatch;
            set
            {
                isMatch = value;

                if (IsMatch && parent != null)
                {
                    //this.isExpanded = true;
                    parent.IsMatch = true;
                }

                NotifyPropertyChanged("IsExpanded");
                NotifyPropertyChanged("IsMatch");
            }
        }

        public bool IsNotRoot => this != RootNode;

        public int Depth => parent == null ? 0 : parent.Depth + 1;

        public bool IsDirectory { get; }

        public MacroFSNode Parent => parent ?? RootNode;

        public bool Equals(MacroFSNode node)
        {
            return FullPath == node.FullPath;
        }

        public ObservableCollection<MacroFSNode> Children
        {
            get
            {
                if (!IsDirectory)
                {
                    return null;
                }

                return children ?? (children = GetChildNodes());
            }
        }

        private ObservableCollection<MacroFSNode> GetChildNodes()
        {
            var files = from childFile in Directory.GetFiles(FullPath)
                        where Path.GetExtension(childFile) == ".js"
                        where childFile != Manager.CurrentMacroPath
                        orderby childFile
                        select childFile;

            var directories = from childDirectory in Directory.GetDirectories(FullPath)
                              orderby childDirectory
                              select childDirectory;

            // Merge files and directories into a collection
            ObservableCollection<MacroFSNode> collection =
                new ObservableCollection<MacroFSNode>(
                    files.Union(directories)
                         .Select((item) => new MacroFSNode(item, this)));

            // Add Current macro at the beginning if this is the root node
            if (this == RootNode)
            {
                collection.Insert(0, new MacroFSNode(Manager.CurrentMacroPath, this));
            }

            return collection;
        }

        public void Delete()
        {
            // If a shortcut is bound to the macro
            if (shortcut > 0)
            {
                // Remove shortcut from shortcut list
                Manager.Shortcuts[shortcut] = string.Empty;
                Manager.Instance.SaveShortcuts(true);
            }

            // Remove macro from collection
            parent.children.Remove(this);

            // Unmonitor the file
            //FileChangeMonitor.Instance.UnmonitorFileSystemEntry(this.FullPath, this.IsDirectory);
        }

        public void EnableEdit()
        {
            IsEditable = true;
        }

        public void DisableEdit()
        {
            IsEditable = false;
        }

        public static void EnableSearch()
        {
            // Set Searching to true
            Searching = true;

            // And then notify all node that their IsMatch property might have changed
            NotifyAllNode(RootNode, "IsMatch");
        }

        public static void DisableSearch()
        {
            // Set Searching to true
            Searching = false;

            UnmatchAllNodes(RootNode);
        }

        /// <summary>
        /// Finds the node with FullPath path in the entire tree 
        /// </summary>
        /// <param name="path"></param>
        /// <returns>MacroFSNode  whose FullPath is path</returns>
        public static MacroFSNode FindNodeFromFullPath(string path)
        {
            if (RootNode == null)
            {
                return null;
            }

            // Default node if search fails
            MacroFSNode defaultNode = RootNode.Children.Count > 0 ? RootNode.Children[0] : RootNode;

            // Make sure path is a valid string
            if (string.IsNullOrEmpty(path))
            {
                return defaultNode;
            }

            // Split the string at '\'
            string shortenPath = path.Substring(path.IndexOf(@"\Macros"));
            string[] substrings = shortenPath.Split(new char[] { '\\' });

            // Starting from the root,
            MacroFSNode node = RootNode;

            try
            {
                // Go down the tree to find the right node
                // 2 because substrings[0] == "" and substrings[1] is root
                for (int i = 3; i < substrings.Length; i++)
                {
                    node = node.Children.Single(x => x.Name == Path.GetFileNameWithoutExtension(substrings[i]));
                }
            }
            catch (Exception e)
            {
                if (ErrorHandler.IsCriticalException(e))
                {
                    throw;
                }

                // Return default node
                node = defaultNode;
            }

            return node;
        }

        public static MacroFSNode SelectNode(string path)
        {
            // Find node
            MacroFSNode node = FindNodeFromFullPath(path);
            if (node != null)
            {
                // Select it
                node.IsSelected = true;
            }

            return node;
        }

        public static void RefreshTree()
        {
            MacroFSNode root = RootNode;
            RefreshTree(root);
        }

        private void AfterRefresh(MacroFSNode root, string selectedPath, HashSet<string> dirs)
        {
            // Set IsEnabled for each folders
            root.SetIsExpanded(root, dirs);

            // Selecte the previously selected macro
            MacroFSNode selected = FindNodeFromFullPath(selectedPath);
            selected.IsSelected = true;

            // Notify change
            root.NotifyPropertyChanged("Children");
        }

        public static void RefreshTree(MacroFSNode root)
        {
            MacroFSNode selected = MacrosControl.Current.MacroTreeView.SelectedItem as MacroFSNode;

            // Make a copy of the hashset
            HashSet<string> dirs = new HashSet<string>(enabledDirectories);

            // Clear enableDirectories
            enabledDirectories.Clear();

            // Retrieve children in a background thread
            //Task.Run(() => root.children = root.GetChildNodes())
            //    .ContinueWith(_ => root.AfterRefresh(root, selected.FullPath, dirs), TaskScheduler.FromCurrentSynchronizationContext());
            root.children = root.GetChildNodes();
            root.AfterRefresh(root, selected.FullPath, dirs);
        }

        public static void CollapseAllNodes(MacroFSNode root)
        {
            if (root.Children != null)
            {
                foreach (var child in root.Children)
                {
                    child.IsExpanded = false;
                    CollapseAllNodes(child);
                }
            }
        }

        public static void UnmatchAllNodes(MacroFSNode root)
        {
            root.isMatch = false;
            root.NotifyPropertyChanged("IsMatch");
            root.NotifyPropertyChanged("IsExpanded");

            if (root.Children != null)
            {
                foreach (var child in root.Children)
                {
                    UnmatchAllNodes(child);
                }
            }
        }

        /// <summary>
        /// Expands all the node marked as expanded in <paramref name="enabledDirs"/>.
        /// </summary>
        /// <param name="node">Tree rooted at node.</param>
        /// <param name="enabledDirs">Hash set containing the enabled dirs.</param>
        private void SetIsExpanded(MacroFSNode node, HashSet<string> enabledDirs)
        {
            node.IsExpanded = true;

            // OPTIMIZATION IDEA instead of iterating over the children, iterate over the enableDirs
            if (node.Children.Count > 0 && enabledDirs.Count > 0)
            {
                foreach (var item in node.children)
                {
                    if (item.IsDirectory && enabledDirs.Remove(item.FullPath))
                    {
                        // Set IsExpanded
                        item.IsExpanded = true;

                        // Recursion on children
                        SetIsExpanded(item, enabledDirs);
                    }
                }
            }
        }

        private void NotifyPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            handler?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // Notifies all the nodes of the tree rooted at 'node'
        private static void NotifyAllNode(MacroFSNode root, string property)
        {
            root.NotifyPropertyChanged(property);

            if (root.Children == null) return;

            foreach (var child in root.Children)
            {
                NotifyAllNode(child, property);
            }
        }
    }
}