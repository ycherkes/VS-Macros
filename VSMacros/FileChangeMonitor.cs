//-----------------------------------------------------------------------
// <copyright file="FileChangeMonitor.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using VSMacros.Models;

namespace VSMacros
{
    class FileChangeMonitor : IVsFileChangeEvents
    {
        private static FileChangeMonitor instance;

        private IVsFileChangeEx fileChangeService;
        private Dictionary<string, uint> cookies;

        public static FileChangeMonitor Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new FileChangeMonitor(VSMacrosPackage.Current);
                }

                return instance;
            }
        }

        public IVsFileChangeEx FileChangeService => fileChangeService;

        private FileChangeMonitor(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            IVsFileChangeEx fileChangeService = serviceProvider.GetService(typeof(SVsFileChangeEx)) as IVsFileChangeEx;
            this.fileChangeService = fileChangeService;

            cookies = new Dictionary<string, uint>();
        }

        #region Monitor and Unmonitor

        /// <summary>
        /// Enables the monitor to receive notifications on changes of a folder.
        /// </summary>
        /// <param name="path">
        /// Folder to monitor.
        /// </param>
        private void MonitorFolder(string path)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            uint cookie;

            FileChangeService.AdviseDirChange(
                path,
                Convert.ToInt32(false),  // Do not monitor subdfolders
                this,
                out cookie
                );

            cookies[path] = cookie;
        }

        /// <summary>
        /// Enables the monitor to receive notifications on changes of a file.
        /// </summary>
        /// <param name="path">
        /// Full path to the file to monitor.
        /// </param>
        private void MonitorFile(string path)
        {
            uint cookie;
            ThreadHelper.ThrowIfNotOnUIThread();
            FileChangeService.AdviseFileChange(
                path,
                (uint)(_VSFILECHANGEFLAGS.VSFILECHG_Add | _VSFILECHANGEFLAGS.VSFILECHG_Attr | _VSFILECHANGEFLAGS.VSFILECHG_Del),
                this,
                out cookie);

            cookies[path] = cookie;
        }

        /// <summary>
        /// Enables the monitor to receive notifications on changes of a file or folder.
        /// </summary>
        /// <param name="path">Full path to the file system entry.</param>
        /// <param name="isDirectory">Boolean telling if the file system entry is a directory.</param>
        public void MonitorFileSystemEntry(string path, bool isDirectory)
        {
            if (isDirectory)
            {
                MonitorFolder(path);
            }
            else
            {
                MonitorFile(path);
            }
        }

        /// <summary>
        /// Disables a client from receiving notifications to a directory.
        /// </summary>
        /// <param name="path">
        /// Full path to the directory.
        /// </param>
        private void UnmonitorFolder(string path)
        {
            uint cookie = cookies[path];
            ThreadHelper.ThrowIfNotOnUIThread();
            FileChangeService.UnadviseDirChange(cookie);
        }

        /// <summary>
        /// Disables a client from receiving notifications to a file.
        /// </summary>
        /// <param name="path">
        /// Full path to the file.
        /// </param>
        private void UnmonitorFile(string path)
        {
            uint cookie = cookies[path];
            ThreadHelper.ThrowIfNotOnUIThread();
            FileChangeService.UnadviseFileChange(cookie);
        }

        /// <summary>
        /// Disables the monitor to receive notifications on changes of a file or folder.
        /// </summary>
        /// <param name="path">Full path to the file system entry.</param>
        /// <param name="isDirectory">Boolean telling if the file system entry is a directory.</param>
        public void UnmonitorFileSystemEntry(string path, bool isDirectory)
        {
            if (isDirectory)
            {
                UnmonitorFolder(path);
            }
            else
            {
                UnmonitorFile(path);
            }
        }

        #endregion
        #region IVsFileChangeEvents Members

        /// <summary>
        /// Notifies clients of changes made to a directory.
        /// </summary>
        /// <param name="dir">
        /// Name of the directory that had a change.
        /// </param>
        public int DirectoryChanged(string dir)
        {
            MacroFSNode node = MacroFSNode.FindNodeFromFullPath(dir);
            if (node != null)
            {
                // Refresh tree rooted at the changed directory
                MacroFSNode.RefreshTree(node);
            }

            return VSConstants.S_OK;
        }

        /// <summary>
        /// Notifies clients of changes made to one or more files.
        /// </summary>
        /// <param name="numberOfFilesChanged">
        /// Number of files changed.
        /// </param>
        /// <param name="files">
        /// Array of file names.
        /// </param>
        /// <param name="typesOfChange">
        /// Array of flags indicating the type of changes. <see cref="_VSFILECHANGEFLAGS" />.
        public int FilesChanged(uint numberOfFilesChanged, string[] files, uint[] changeTypes)
        {
            // Go over each file and treat the change appropriately
            for (int i = 0; i < files.Length; i++)
            {
                string path = files[i];
                _VSFILECHANGEFLAGS change = (_VSFILECHANGEFLAGS)changeTypes[i];

                // _VSFILECHANGEFLAGS.VSFILECHG_Add is handled by DirectoryChanged
                // Only handle _VSFILECHANGEFLAGS.VSFILECHG_Del here
                if (change == _VSFILECHANGEFLAGS.VSFILECHG_Del)
                {
                    MacroFSNode node = MacroFSNode.FindNodeFromFullPath(path);
                    if (node != null)
                    {
                        node.Delete();
                    }
                }
            }

            return VSConstants.S_OK;
        }

        #endregion
    }
}
