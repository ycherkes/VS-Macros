﻿//-----------------------------------------------------------------------
// <copyright file="FileChangeMonitor.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Models;

namespace VSMacros
{
    class FileChangeMonitor : IVsFileChangeEvents
    {
        private static FileChangeMonitor instance;

        private Dictionary<string, uint> cookies;

        public static FileChangeMonitor Instance => instance ??
                                                    (instance = new FileChangeMonitor(VSMacrosPackage.Current));

        public IVsFileChangeEx FileChangeService { get; }

        private FileChangeMonitor(IServiceProvider serviceProvider)
        {
            var fileChangeService = serviceProvider.GetService(typeof(SVsFileChangeEx)) as IVsFileChangeEx;
            FileChangeService = fileChangeService;

            cookies = new Dictionary<string, uint>();
        }

        #region Monitor and Unmonitor

        /// <summary>
        /// Enables the monitor to receive notifications on changes of a folder.
        /// </summary>
        /// <param name="path">
        /// Folder to monitor.
        /// </param>
        public void MonitorFolder(string path)
        {
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
        public void MonitorFile(string path)
        {
            uint cookie;

            FileChangeService.AdviseFileChange(
                path,
                (uint)(_VSFILECHANGEFLAGS.VSFILECHG_Add | _VSFILECHANGEFLAGS.VSFILECHG_Attr | _VSFILECHANGEFLAGS.VSFILECHG_Del),
                this,
                out cookie);

            this.cookies[path] = cookie;
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
                this.MonitorFolder(path);
            }
            else
            {
                this.MonitorFile(path);
            }
        }

        /// <summary>
        /// Disables a client from receiving notifications to a directory.
        /// </summary>
        /// <param name="path">
        /// Full path to the directory.
        /// </param>
        public void UnmonitorFolder(string path)
        {
            var cookie = cookies[path];

            FileChangeService.UnadviseDirChange(cookie);
        }

        /// <summary>
        /// Disables a client from receiving notifications to a file.
        /// </summary>
        /// <param name="path">
        /// Full path to the file.
        /// </param>
        public void UnmonitorFile(string path)
        {
            var cookie = cookies[path];

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
            var node = MacroFSNode.FindNodeFromFullPath(dir);
            if(node != null)
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
        /// <param name="changeTypes"/>
        /// Array of flags indicating the type of changes. <see cref="_VSFILECHANGEFLAGS" />.
        public int FilesChanged(uint numberOfFilesChanged, string[] files, uint[] changeTypes)
        {
            // Go over each file and treat the change appropriately
            for ( var i = 0; i < files.Length; i++)
            {
                var path = files[i];
                var change = (_VSFILECHANGEFLAGS)changeTypes[i];

                // _VSFILECHANGEFLAGS.VSFILECHG_Add is handled by DirectoryChanged
                // Only handle _VSFILECHANGEFLAGS.VSFILECHG_Del here
                if (change != _VSFILECHANGEFLAGS.VSFILECHG_Del) continue;
                var node = MacroFSNode.FindNodeFromFullPath(path);
                node?.Delete();
            }

            return VSConstants.S_OK;
        }

        #endregion
    }
}
