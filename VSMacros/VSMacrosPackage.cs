//-----------------------------------------------------------------------
// <copyright file="VSMacrosPackage.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using EnvDTE;
using Microsoft.Internal.VisualStudio.PlatformUI;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using EnvDTE80;
using VSMacros.Engines;
using VSMacros.Interfaces;
using VSMacros.Model;

namespace VSMacros
{
    [Guid(GuidList.GuidVSMacrosPkgString)]
    [ProvideToolWindow(typeof(MacrosToolWindow), Style = VsDockStyle.Tabbed, Window = "3ae79031-e1bc-11d0-8f78-00a0c9110057")]
    public sealed class VSMacrosPackage : AsyncPackage
    {
        private static VSMacrosPackage current;
        public static VSMacrosPackage Current => current ?? (current = new VSMacrosPackage());

        public VSMacrosPackage()
        {
            current = this;
        }

        private void ShowToolWindow(object sender = null, EventArgs e = null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Get the (only) instance of this tool window
            // The last flag is set to true so that if the tool window does not exists it will be created.
            ToolWindowPane window = FindToolWindow(typeof(MacrosToolWindow), 0, true);
            if ((window == null) || (window.Frame == null))
            {
                throw new NotSupportedException(Resources.CannotCreateWindow);
            }
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        private string macroDirectory;
        public string MacroDirectory => macroDirectory ?? (macroDirectory = Path.Combine(UserLocalDataPath, "Macros"));

        private string assemblyDirectory;
        public string AssemblyDirectory
        {
            get
            {
                if (assemblyDirectory == default(string))
                {
                    assemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                }

                return assemblyDirectory;
            }
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members
        private BitmapImage startIcon;
        private BitmapImage playbackIcon;
        private BitmapImage stopIcon;
        private string commonPath;
        private List<CommandBarButton> imageButtons;
        private IVsStatusbar statusBar;
        private static RecorderDataModel dataModel;

        internal static RecorderDataModel DataModel => dataModel ?? (dataModel = new RecorderDataModel());

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            await base.InitializeAsync(cancellationToken, progress);

            ((IServiceContainer)this).AddService(typeof(IRecorder), (serviceContainer, type) => new Recorder(this), promote: true);
            statusBar = (IVsStatusbar)await GetServiceAsync(typeof(SVsStatusbar));

            // Add our command handlers for the menu
            OleMenuCommandService mcs = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the tool window
                mcs.AddCommand(new MenuCommand(
                   ShowToolWindow,
                   new CommandID(GuidList.GuidVSMacrosCmdSet, (int)PkgCmdIDList.CmdIdMacroExplorer)));

                // Create the command for start recording
                CommandID recordCommandID = new CommandID(GuidList.GuidVSMacrosCmdSet, (int)PkgCmdIDList.CmdIdRecord);
                OleMenuCommand recordMenuItem = new OleMenuCommand(Record, recordCommandID);
                recordMenuItem.BeforeQueryStatus += Record_OnBeforeQueryStatus;
                mcs.AddCommand(recordMenuItem);

                // Create the command for playback
                mcs.AddCommand(new MenuCommand(
                    Playback,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdPlayback)));

                // Create the command for playback multiple times
                mcs.AddCommand(new MenuCommand(
                    PlaybackMultipleTimes,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdPlaybackMultipleTimes)));

                // Create the command for save current macro
                mcs.AddCommand(new MenuCommand(
                    SaveCurrent,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdSaveTemporaryMacro)));

                // Create the command to playback bounded macros
                mcs.AddCommand(new MenuCommand(PlaybackCommand1, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand1)));
                mcs.AddCommand(new MenuCommand(PlaybackCommand2, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand2)));
                mcs.AddCommand(new MenuCommand(PlaybackCommand3, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand3)));
                mcs.AddCommand(new MenuCommand(PlaybackCommand4, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand4)));
                mcs.AddCommand(new MenuCommand(PlaybackCommand5, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand5)));
                mcs.AddCommand(new MenuCommand(PlaybackCommand6, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand6)));
                mcs.AddCommand(new MenuCommand(PlaybackCommand7, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand7)));
                mcs.AddCommand(new MenuCommand(PlaybackCommand8, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand8)));
                mcs.AddCommand(new MenuCommand(PlaybackCommand9, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand9)));
            }
        }
        #endregion

        /////////////////////////////////////////////////////////////////////////////
        // Command Handlers
        #region Command Handlers

        private void Record(object sender, EventArgs arguments)
        {
            IRecorderPrivate macroRecorder = (IRecorderPrivate)GetService(typeof(IRecorder));
            if (!macroRecorder.IsRecording)
            {
                Manager.Instance.StartRecording();

                StatusBarChange(Resources.StatusBarRecordingText, 1);
                ChangeMenuIcons(StopIcon, 0);
                UpdateButtonsForRecording(true);
            }
            else
            {
                Manager.Instance.StopRecording();

                StatusBarChange(Resources.StatusBarReadyText, 0);
                ChangeMenuIcons(StartIcon, 0);
                UpdateButtonsForRecording(false);
            }
        }

        private void Playback(object sender, EventArgs arguments)
        {
            if (Manager.Instance.executor == null || !Manager.Instance.executor.IsEngineRunning)
            {
                //this.UpdateButtonsForPlayback(true);
            }
            else
            {
                UpdateButtonsForPlayback(false);
            }

            Manager.Instance.Playback(string.Empty);
        }

        private void PlaybackMultipleTimes(object sender, EventArgs arguments)
        {
            if (Manager.Instance.executor == null || !Manager.Instance.executor.IsEngineRunning)
            {
                //this.UpdateButtonsForPlayback(true);
            }
            else
            {
                UpdateButtonsForPlayback(false);
            }

            Manager.Instance.PlaybackMultipleTimes(string.Empty);
        }

        private static void SaveCurrent(object sender, EventArgs arguments)
        {
            Manager.Instance.SaveCurrent();
        }

        private static void PlaybackCommand1(object sender, EventArgs arguments) { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(1); }
        private static void PlaybackCommand2(object sender, EventArgs arguments) { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(2); }
        private static void PlaybackCommand3(object sender, EventArgs arguments) { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(3); }
        private static void PlaybackCommand4(object sender, EventArgs arguments) { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(4); }
        private static void PlaybackCommand5(object sender, EventArgs arguments) { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(5); }
        private static void PlaybackCommand6(object sender, EventArgs arguments) { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(6); }
        private static void PlaybackCommand7(object sender, EventArgs arguments) { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(7); }
        private static void PlaybackCommand8(object sender, EventArgs arguments) { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(8); }
        private static void PlaybackCommand9(object sender, EventArgs arguments) { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(9); }

        #endregion

        #region Status Bar & Menu Icons
        public void ChangeMenuIcons(BitmapSource icon, int commandNumber)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // commandNumber is 0 for Recording, 1 for Playback and 2 for Playback Multiple Times         
            try
            {
                if (ImageButtons[commandNumber] == null) return;
                // Change icon in menu
                ImageButtons[commandNumber].Picture = (stdole.StdPicture)ImageHelper.IPictureFromBitmapSource(icon);

                if (ImageButtons.Count > 3)
                {
                    // Change icon in toolbar
                    ImageButtons[commandNumber + 3].Picture = (stdole.StdPicture)ImageHelper.IPictureFromBitmapSource(icon);
                }
            }
            catch (ObjectDisposedException)
            {
                // Do nothing since the removed button does not need to change its image;
            }
        }

        internal void StatusBarChange(string status, int animation)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            statusBar.Clear();
            statusBar.SetText(status);
        }

        internal List<CommandBarButton> ImageButtons
        {
            get
            {
                if (imageButtons != null) return imageButtons;

                imageButtons = new List<CommandBarButton>();
                AddMenuButton();
                return imageButtons;
            }
        }

        private void AddMenuButton()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE2 dte = (DTE2)GetService(typeof(SDTE));
            CommandBar mainMenu = ((CommandBars)dte.CommandBars)["MenuBar"];
            CommandBarPopup toolMenu = (CommandBarPopup)mainMenu.Controls["Tools"];
            CommandBarPopup macroMenu = (CommandBarPopup)toolMenu.Controls["Macros"];
            if (macroMenu != null)
            {
                try
                {
                    List<CommandBarButton> buttons = new List<CommandBarButton>()
                    {
                        (CommandBarButton)macroMenu.Controls["Start Recording"],
                        (CommandBarButton)macroMenu.Controls["Playback"],
                        (CommandBarButton)macroMenu.Controls["Playback Multiple Times"]
                    };

                    foreach (var item in buttons)
                    {
                        if (item != null)
                        {
                            imageButtons.Add(item);
                        }
                    }
                }
                catch (ArgumentException)
                {
                    // nothing to do
                }
            }
        }

        private void UpdateButtonsForRecording(bool isRecording)
        {
            EnableMyCommand(PkgCmdIDList.CmdIdPlayback, !isRecording);
            EnableMyCommand(PkgCmdIDList.CmdIdPlaybackMultipleTimes, !isRecording);
            UpdateCommonButtons(!isRecording);
        }

        public void UpdateButtonsForPlayback(bool goingToPlay)
        {
            EnableMyCommand(PkgCmdIDList.CmdIdRecord, !goingToPlay);
            EnableMyCommand(PkgCmdIDList.CmdIdPlaybackMultipleTimes, !goingToPlay);
            UpdateCommonButtons(!goingToPlay);

            ChangeMenuIcons(goingToPlay ? StopIcon : PlaybackIcon, 1);
        }

        private void UpdateCommonButtons(bool enable)
        {
            EnableMyCommand(PkgCmdIDList.CmdIdSaveTemporaryMacro, enable);
            EnableMyCommand(PkgCmdIDList.CmdIdRefresh, enable);
            EnableMyCommand(PkgCmdIDList.CmdIdOpenDirectory, enable);
        }

        private bool EnableMyCommand(int cmdId, bool enableCmd)
        {
            bool cmdUpdated = false;
            MenuCommand mc;
            using (var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService)
            {
                var newCmdId = new CommandID(GuidList.GuidVSMacrosCmdSet, cmdId);
                mc = mcs.FindCommand(newCmdId);
            }

            if (mc != null)
            {
                mc.Enabled = enableCmd;
                cmdUpdated = true;
            }
            return cmdUpdated;
        }

        internal void ClearStatusBar()
        {
            StatusBarChange(Resources.StatusBarReadyText, 0);
        }

        private BitmapSource StartIcon => startIcon ?? (startIcon = new BitmapImage(new Uri(Path.Combine(CommonPath, "RecordRound.png"))));

        private BitmapSource PlaybackIcon =>
            playbackIcon ??
            (playbackIcon = new BitmapImage(new Uri(Path.Combine(CommonPath, "PlaybackIcon.png"))));

        internal BitmapSource StopIcon => stopIcon ?? (stopIcon = new BitmapImage(new Uri(Path.Combine(CommonPath, "StopIcon.png"))));

        #endregion

        private string CommonPath =>
            commonPath ?? (commonPath =
                Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources"));

        protected override int QueryClose(out bool canClose)
        {
            IRecorderPrivate macroRecorder = (IRecorderPrivate)GetService(typeof(IRecorder));
            if (macroRecorder.IsRecording)
            {
                string message = Resources.ExitMessage;
                string caption = Resources.ExitCaption;
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;

                // Displays the MessageBox.
                var result = MessageBox.Show(message, caption, buttons);
                canClose = result == System.Windows.Forms.DialogResult.Yes;
            }
            else
            {
                canClose = true;
            }

            // Close manager
            Manager.Instance.Close();

            if (Executor.Job != null)
            {
                Executor.Job.Close();
            }

            return VSConstants.S_OK;
        }

        private void Record_OnBeforeQueryStatus(object sender, EventArgs e)
        {
            var recordCommand = sender as OleMenuCommand;

            if (recordCommand == null) return;

            IRecorderPrivate macroRecorder = (IRecorderPrivate)GetService(typeof(IRecorder));
            recordCommand.Text = macroRecorder.IsRecording ? Resources.MenuTextRecording : Resources.MenuTextNormal;
        }
    }
}
