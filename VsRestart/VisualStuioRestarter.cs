using System;
using System.Diagnostics;
using System.Linq;
using System.Text;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using MidnightDevelopers.VisualStudio.VsRestart.Arguments;
using Process = System.Diagnostics.Process;

namespace MidnightDevelopers.VisualStudio.VsRestart
{
    internal class VisualStuioRestarter
    {
        internal void Restart(DTE dte, bool elevated)
        {
            var currentProcess = Process.GetCurrentProcess();

            var parser = new ArgumentParser(dte.CommandLineArguments);

            var builder = new RestartProcessBuilder()
                .WithDevenv(currentProcess.MainModule.FileName)
                .WithArguments(parser.GetArguments());

            var openedItem = GetOpenedItem(dte);
            if (openedItem != OpenedItem.None)
            {
                if (openedItem.IsSolution)
                {
                    builder.WithSolution(openedItem.Name);
                }
                else
                {
                    builder.WithProject(openedItem.Name);
                }
            }

            if (elevated)
            {
                builder.WithElevatedPermission();
            }

            const string commandName = "File.Exit";
            var closeCommand = dte.Commands.Item(commandName);

            CommandEvents closeCommandEvents = null;
            if (closeCommand != null)
            {
                closeCommandEvents = dte.Events.CommandEvents[closeCommand.Guid, closeCommand.ID];
            }

            // Install the handler
            var handler = new VisualStudioEventHandler(dte.Events.DTEEvents, closeCommandEvents, builder.Build());

            if (closeCommand != null && closeCommand.IsAvailable)
            {
                // if the Exit commad is present, execute it with all gracefulls dialogs by VS
                dte.ExecuteCommand(commandName);
            }
            else
            {
                // Close brutally
                dte.Quit();
            }
        }

        private static OpenedItem GetOpenedItem(DTE dte)
        {
            if (dte.Solution != null && dte.Solution.IsOpen)
            {
                if (string.IsNullOrEmpty(dte.Solution.FullName))
                {
                    Array activeProjects = (Array)dte.ActiveSolutionProjects;
                    if (activeProjects != null && activeProjects.Length > 0)
                    {
                        var currentOpenedProject = (Project)activeProjects.GetValue(0);
                        if (currentOpenedProject != null)
                        {
                            return new OpenedItem(currentOpenedProject.FullName, false);
                        }
                    }
                }
                else
                {
                    return new OpenedItem(dte.Solution.FullName, true);
                }
            }

            return OpenedItem.None;
        }

        private enum ProcessStartResult
        {
            Ok,
            AuthDenied,
            Exception,
        }

        private class RestartProcessBuilder
        {
            private string _solutionFile;
            private string _devenv;
            private ArgumentTokenCollection _arguments;
            private string _projectFile;
            private string _verb = null;

            public RestartProcessBuilder WithSolution(string solutionFile)
            {
                _solutionFile = solutionFile;
                return this;
            }

            public RestartProcessBuilder WithArguments(ArgumentTokenCollection arguments)
            {
                _arguments = arguments;
                return this;
            }

            public RestartProcessBuilder WithDevenv(string devenv)
            {
                _devenv = devenv;
                return this;
            }

            public RestartProcessBuilder WithProject(string projectFile)
            {
                _projectFile = projectFile;
                return this;
            }

            public RestartProcessBuilder WithElevatedPermission()
            {
                _verb = "runas";
                return this;
            }

            public ProcessStartInfo Build()
            {
                return new ProcessStartInfo
                {
                    FileName = _devenv,
                    ErrorDialog = true,
                    UseShellExecute = true,
                    Verb = _verb,
                    Arguments = BuildArguments(),
                };
            }

            private string BuildArguments()
            {
                if (!string.IsNullOrEmpty(_solutionFile))
                {
                    if (_arguments.OfType<SolutionArgumentToken>().Any())
                    {
                        _arguments.Replace<SolutionArgumentToken>(new SolutionArgumentToken(Quote(_solutionFile)));
                    }
                    else if (_arguments.OfType<ProjectArgumentToken>().Any())
                    {
                        _arguments.Replace<ProjectArgumentToken>(new SolutionArgumentToken(Quote(_solutionFile)));
                    }
                    else
                    {
                        _arguments.Add(new SolutionArgumentToken(Quote(_solutionFile)));
                    }
                }

                if (!string.IsNullOrEmpty(_projectFile))
                {
                    if (_arguments.OfType<SolutionArgumentToken>().Any())
                    {
                        _arguments.Replace<SolutionArgumentToken>(new ProjectArgumentToken(Quote(_projectFile)));
                    }
                    else if (_arguments.OfType<ProjectArgumentToken>().Any())
                    {
                        _arguments.Replace<ProjectArgumentToken>(new ProjectArgumentToken(Quote(_projectFile)));
                    }
                    else
                    {
                        _arguments.Add(new ProjectArgumentToken(Quote(_projectFile)));
                    }
                }

                string escapedArguments = _arguments.ToString()
                    .ReplaceSmart(Quote(_devenv), string.Empty);

                return escapedArguments;
            }

            private string Quote(string input)
            {
                return string.Format("\"{0}\"", input);
            }
        }

        private class OpenedItem
        {
            public static readonly OpenedItem None = new OpenedItem(null, false);

            public OpenedItem(string name, bool isSolution)
            {
                Name = name;
                IsSolution = isSolution;
            }

            public string Name { get; private set; }

            public bool IsSolution { get; set; }
        }

        private class VisualStudioEventHandler
        {
            private readonly ProcessStartInfo _startInfo;
            private readonly CommandEvents _closeCommandEvents;
            private readonly DTEEvents _dTEEvents;

            public VisualStudioEventHandler(DTEEvents dTEEvents, CommandEvents closeCommandEvents, ProcessStartInfo processStartInfo)
            {
                _dTEEvents = dTEEvents;
                _closeCommandEvents = closeCommandEvents;
                _startInfo = processStartInfo;

                dTEEvents.OnBeginShutdown += DTEEvents_OnBeginShutdown;
                if (closeCommandEvents != null)
                {
                    closeCommandEvents.AfterExecute += CommandEvents_AfterExecute;
                }
            }

            void CommandEvents_AfterExecute(string guid, int id, object customIn, object customOut)
            {
                _dTEEvents.OnBeginShutdown -= DTEEvents_OnBeginShutdown;
                if (_closeCommandEvents != null)
                {
                    _closeCommandEvents.AfterExecute -= CommandEvents_AfterExecute;
                }
            }

            void DTEEvents_OnBeginShutdown()
            {
                if (StartProcessSafe(_startInfo, DisplayError) == ProcessStartResult.Ok)
                {
                    //currentProcess.Kill();
                }
            }

            private static ProcessStartResult StartProcessSafe(ProcessStartInfo startInfo, Action<Exception, ProcessStartResult> exceptionHandler)
            {
                ProcessStartResult result = ProcessStartResult.Ok;

                try
                {
                    Process.Start(startInfo);
                }
                catch (Exception ex)
                {
                    result = ProcessStartResult.Exception;
                    var winex = ex as System.ComponentModel.Win32Exception;

                    // User has denied auth through UAC
                    if (winex != null && winex.NativeErrorCode == 1223)
                    {
                        result = ProcessStartResult.AuthDenied;
                    }

                    exceptionHandler(ex, result);
                }

                return result;
            }

            private static void DisplayError(Exception ex, ProcessStartResult status)
            {
                IVsOutputWindowPane outputPane = Package.GetGlobalService(typeof(SVsGeneralOutputWindowPane)) as IVsOutputWindowPane;

                outputPane.Activate();

                if (status == ProcessStartResult.AuthDenied)
                {
                    outputPane.OutputString("Visual Studio restart operation was cancelled by the user." + Environment.NewLine);
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("An exceptions has been trown while trying to start an elevated Visual Studio, see details below.");
                    sb.AppendLine(ex.ToString());

                    string diagnostics = sb.ToString();

                    outputPane.OutputString(diagnostics);
                    IVsActivityLog log = Package.GetGlobalService(typeof(SVsActivityLog)) as IVsActivityLog;
                    log.LogEntry((uint)__ACTIVITYLOG_ENTRYTYPE.ALE_ERROR, "MidnightDevelopers.VsRestarter", diagnostics);
                }

                //EnvDTE.OutputWindow.OutputWindow.Parent.Activate();
            }
        }
    }
}