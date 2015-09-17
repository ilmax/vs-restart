using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Process = System.Diagnostics.Process;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;

namespace MidnightDevelopers.VisualStudio.VsRestart
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.VsRestarterPackageId)]
    [ProvideAutoLoad("adfc4e64-0397-11d1-9f4e-00a0c911004f")]
    public sealed class VsRestartPackage : Package
    {
        private DTE __DTE;
        private DteInitializer _dteInitializer;

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            InitializeDTE();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {

                CommandID restartCommandSingleID = new CommandID(GuidList.TopLevelMenuGuid, (int)MenuId.Restart);
                OleMenuCommand restartSingleMenuItem = new OleMenuCommand(RestartMenuItemCallback, restartCommandSingleID);
                mcs.AddCommand(restartSingleMenuItem);

                restartSingleMenuItem.BeforeQueryStatus += OnBeforeQueryStatusSingle;

                // Create the command for the menu item.
                CommandID restartElevatedCommandID = new CommandID(GuidList.RestartElevatedGroupGuid, (int)MenuId.RestartAsAdmin);
                OleMenuCommand restartElevatedMenuItem = new OleMenuCommand(RestartMenuItemCallback, restartElevatedCommandID);
                mcs.AddCommand(restartElevatedMenuItem);

                restartElevatedMenuItem.BeforeQueryStatus += OnBeforeQueryStatusGroup;

                CommandID restartCommandID = new CommandID(GuidList.RestartElevatedGroupGuid, (int)MenuId.Restart);
                OleMenuCommand restartMenuItem = new OleMenuCommand(RestartMenuItemCallback, restartCommandID);
                mcs.AddCommand(restartMenuItem);

                restartMenuItem.BeforeQueryStatus += OnBeforeQueryStatusGroup;
            }
        }

        private void InitializeDTE()
        {
            try
            {
                __DTE = (DTE)GetService(typeof(DTE));
            }
            catch (Exception)
            {
                __DTE = null;
            }

            if (__DTE == null)
            {
                IVsShell shellService = (IVsShell)this.GetService(typeof(IVsShell));
                _dteInitializer = new DteInitializer(shellService, InitializeDTE);
            }
            else
            {
                _dteInitializer = null;
            }
        }

        private void OnBeforeQueryStatusGroup(object sender, EventArgs e)
        {
            // I don't need this next line since this is a lambda.
            // But I just wanted to show that sender is the OleMenuCommand.
            OleMenuCommand item = (OleMenuCommand)sender;
            if (ElevationChecker.CanCheckElevation)
            {
                item.Visible = !ElevationChecker.IsElevated(Process.GetCurrentProcess().Handle);
            }
        }

        private void OnBeforeQueryStatusSingle(object sender, EventArgs e)
        {
            // I don't need this next line since this is a lambda.
            // But I just wanted to show that sender is the OleMenuCommand.
            OleMenuCommand item = (OleMenuCommand)sender;
            if (ElevationChecker.CanCheckElevation)
            {
                item.Visible = ElevationChecker.IsElevated(Process.GetCurrentProcess().Handle);
            }
        }

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void RestartMenuItemCallback(object sender, EventArgs e)
        {
            var dte = __DTE;

            if (dte == null)
            {
                // Show some error message and return
                return;
            }

            Debug.Assert(dte != null);

            bool elevated = ((OleMenuCommand)sender).CommandID.ID == MenuId.RestartAsAdmin;

            new VisualStuioRestarter().Restart(dte, elevated);
        }

        private bool CanClose()
        {

            bool a;
            if (QueryClose(out a) == VSConstants.S_OK)
            {
                return a;
            }

            return false;
        }
    }

    // Courtesy of http://www.mztools.com/articles/2013/MZ2013029.aspx
    internal class DteInitializer : IVsShellPropertyEvents
    {
        private readonly IVsShell _shellService;
        private uint _cookie;
        private readonly Action _callback;

        internal DteInitializer(IVsShell shellService, Action callback)
        {
            int hr;

            _shellService = shellService;
            _callback = callback;

            // Set an event handler to detect when the IDE is fully initialized
            hr = _shellService.AdviseShellPropertyChanges(this, out _cookie);

            ErrorHandler.ThrowOnFailure(hr);
        }

        int IVsShellPropertyEvents.OnShellPropertyChange(int propid, object var)
        {
            if (propid == (int)__VSSPROPID.VSSPROPID_Zombie)
            {
                var isZombie = (bool)var;

                if (!isZombie)
                {
                    // Release the event handler to detect when the IDE is fully initialized
                    var hr = _shellService.UnadviseShellPropertyChanges(_cookie);

                    ErrorHandler.ThrowOnFailure(hr);

                    _cookie = 0;

                    _callback();
                }
            }

            return VSConstants.S_OK;
        }
    }
}
