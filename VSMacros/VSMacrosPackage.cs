//-----------------------------------------------------------------------
// <copyright file="VSMacrosPackage.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

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

namespace VSMacros
{
    [Guid(GuidList.GuidVSMacrosPkgString)]
    [ProvideToolWindow(typeof(MacrosToolWindow), Style = VsDockStyle.Tabbed, Window = "3ae79031-e1bc-11d0-8f78-00a0c9110057")]
    public sealed class VSMacrosPackage : AsyncPackage
    {
        private static VSMacrosPackage _current;
        public static VSMacrosPackage Current => _current ?? (_current = new VSMacrosPackage());

        public VSMacrosPackage()
        {
            _current = this;
        }

        public async Task ShowToolWindowAsync()
        {
            await JoinableTaskFactory.RunAsync(async delegate
            {
                ToolWindowPane window = await ShowToolWindowAsync(typeof(MacrosToolWindow), 0, true, DisposalToken);
                if (null == window || null == window.Frame)
                {
                    throw new NotSupportedException("Cannot create tool window");
                }

                await JoinableTaskFactory.SwitchToMainThreadAsync();
                IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
                ErrorHandler.ThrowOnFailure(windowFrame.Show());
            });
        }

        private string macroDirectory;
        public string MacroDirectory => macroDirectory ?? (macroDirectory = Path.Combine(UserLocalDataPath, "Macros"));

        private string assemblyDirectory;

        public string AssemblyDirectory =>
            assemblyDirectory ??
            (assemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation

        #region Package Members

        private BitmapImage startIcon;
        private BitmapImage playbackIcon;
        private BitmapImage stopIcon;
        private string commonPath;
        private List<CommandBarButton> imageButtons;
        private IVsStatusbar statusBar;
        private Dictionary<int, MenuCommand> menuCommands = new Dictionary<int, MenuCommand>();

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            await base.InitializeAsync(cancellationToken, progress);

            ((IServiceContainer)this).AddService(typeof(IRecorder), (serviceContainer, type) => new Recorder(this), promote: true);
            statusBar = (IVsStatusbar)await GetServiceAsync(typeof(SVsStatusbar));

            // Create the command for start recording
            CommandID recordCommandId = new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdRecord);
            OleMenuCommand recordMenuItem = new OleMenuCommand(Record, recordCommandId);
            recordMenuItem.BeforeQueryStatus += Record_OnBeforeQueryStatus;
            menuCommands.Add(PkgCmdIDList.CmdIdRecord, recordMenuItem);

            menuCommands = new Dictionary<int, MenuCommand>
            {
                {
                    // Create the command for the tool window
                    PkgCmdIDList.CmdIdMacroExplorer,
                    new MenuCommand((s, e) => ThreadHelper.JoinableTaskFactory.Run(async () => await ShowToolWindowAsync()),
                        new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdMacroExplorer))
                },
                {
                    // Create the command for start recording
                    PkgCmdIDList.CmdIdRecord,
                    recordMenuItem
                },
                {
                    // Create the command for playback
                    PkgCmdIDList.CmdIdPlayback,
                    new MenuCommand(
                    Playback,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdPlayback))
                },
                {
                    // Create the command for playback multiple times
                    PkgCmdIDList.CmdIdPlaybackMultipleTimes, new MenuCommand(
                        PlaybackMultipleTimes,
                        new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdPlaybackMultipleTimes))
                },
                {
                    // Create the command for save current macro
                    PkgCmdIDList.CmdIdSaveTemporaryMacro, new MenuCommand(
                        SaveCurrent,
                        new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdSaveTemporaryMacro))
                },
                // Create the command to playback bounded macros
                { PkgCmdIDList.CmdIdCommand1, new MenuCommand(PlaybackCommand1, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand1)) },
                { PkgCmdIDList.CmdIdCommand2, new MenuCommand(PlaybackCommand2, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand2)) },
                { PkgCmdIDList.CmdIdCommand3, new MenuCommand(PlaybackCommand3, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand3)) },
                { PkgCmdIDList.CmdIdCommand4, new MenuCommand(PlaybackCommand4, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand4)) },
                { PkgCmdIDList.CmdIdCommand5, new MenuCommand(PlaybackCommand5, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand5)) },
                { PkgCmdIDList.CmdIdCommand6, new MenuCommand(PlaybackCommand6, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand6)) },
                { PkgCmdIDList.CmdIdCommand7, new MenuCommand(PlaybackCommand7, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand7)) },
                { PkgCmdIDList.CmdIdCommand8, new MenuCommand(PlaybackCommand8, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand8)) },
                { PkgCmdIDList.CmdIdCommand9, new MenuCommand(PlaybackCommand9, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand9)) }
            };

            // Add our command handlers for the menu
            if (await GetServiceAsync(typeof(IMenuCommandService)) is OleMenuCommandService mcs)
            {
                foreach (var command in menuCommands.Values)
                {
                    mcs.AddCommand(command);
                }
            }
        }

        #endregion Package Members

        /////////////////////////////////////////////////////////////////////////////
        // Command Handlers

        #region Command Handlers

        private void Record(object sender, EventArgs arguments)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            IRecorderPrivate macroRecorder = (IRecorderPrivate)GetService(typeof(IRecorder));
            if (!macroRecorder.IsRecording)
            {
                Manager.Instance.StartRecording();

                StatusBarChange(Resources.StatusBarRecordingText);
                ChangeMenuIcons(StopIcon, 0);
                UpdateButtonsForRecording(true);
            }
            else
            {
                Manager.Instance.StopRecording();

                StatusBarChange(Resources.StatusBarReadyText);
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
            if (Manager.Instance.executor != null && Manager.Instance.executor.IsEngineRunning)
            {
                UpdateButtonsForPlayback(false);
            }

            Manager.Instance.PlaybackMultipleTimes();
        }

        private static void SaveCurrent(object sender, EventArgs arguments)
        {
            Manager.Instance.SaveCurrent();
        }

        private static void PlaybackCommand1(object sender, EventArgs arguments)
        { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(1); }

        private static void PlaybackCommand2(object sender, EventArgs arguments)
        { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(2); }

        private static void PlaybackCommand3(object sender, EventArgs arguments)
        { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(3); }

        private static void PlaybackCommand4(object sender, EventArgs arguments)
        { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(4); }

        private static void PlaybackCommand5(object sender, EventArgs arguments)
        { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(5); }

        private static void PlaybackCommand6(object sender, EventArgs arguments)
        { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(6); }

        private static void PlaybackCommand7(object sender, EventArgs arguments)
        { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(7); }

        private static void PlaybackCommand8(object sender, EventArgs arguments)
        { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(8); }

        private static void PlaybackCommand9(object sender, EventArgs arguments)
        { ThreadHelper.ThrowIfNotOnUIThread(); Manager.Instance.PlaybackCommand(9); }

        #endregion Command Handlers

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

        internal void StatusBarChange(string status)
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

        private void EnableMyCommand(int cmdId, bool enableCmd)
        {
            if (menuCommands.TryGetValue(cmdId, out var command))
            {
                command.Enabled = enableCmd;
            }
        }

        internal void ClearStatusBar()
        {
            StatusBarChange(Resources.StatusBarReadyText);
        }

        private BitmapSource StartIcon => startIcon ?? (startIcon = new BitmapImage(new Uri(Path.Combine(CommonPath, "RecordRound.png"))));

        private BitmapSource PlaybackIcon =>
            playbackIcon ??
            (playbackIcon = new BitmapImage(new Uri(Path.Combine(CommonPath, "PlaybackIcon.png"))));

        internal BitmapSource StopIcon => stopIcon ?? (stopIcon = new BitmapImage(new Uri(Path.Combine(CommonPath, "StopIcon.png"))));

        #endregion Status Bar & Menu Icons

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