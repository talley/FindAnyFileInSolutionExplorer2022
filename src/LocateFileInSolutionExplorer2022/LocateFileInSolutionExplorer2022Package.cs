using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using static Microsoft.VisualStudio.VSConstants;
using Task = System.Threading.Tasks.Task;

namespace LocateFileInSolutionExplorer2022
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks[Guid(FindInSolutionExplorerConstants.guidFindInSolutionExplorerPackageString)]

    [Guid(FindInSolutionExplorerConstants.guidFindInSolutionExplorerPackageString)]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource(1000, 1)]
    [ProvideAutoLoad(UICONTEXT.SolutionExists_string)]
    internal  class LocateFileInSolutionExplorer2022Package : Package
    {
         /// <summary>
        /// LocateFileInSolutionExplorer2022Package GUID string.
        /// </summary>
        public const string PackageGuidString = "115ab1f8-846c-4941-91d3-11fa8e04dff0";

        private static LocateFileInSolutionExplorer2022Package _instance;

        private OleMenuCommand _command;
        public LocateFileInSolutionExplorer2022Package()
        {
            _instance = this;
        }

        public static LocateFileInSolutionExplorer2022Package Instance
        {
            get
            {
                return _instance;
            }
        }

        public SVsServiceProvider ServiceProvider
        {
            get
            {
                return new VsServiceProviderWrapper(this);
            }
        }

        public EnvDTE80.DTE2 ApplicationObject
        {
            get
            {
                return ServiceProvider.GetService(typeof(EnvDTE._DTE)) as EnvDTE80.DTE2;
            }
        }

        protected override void Initialize()
        {
            base.Initialize();
            IMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as IMenuCommandService;

            CommandID id = new CommandID(FindInSolutionExplorerConstants.guidFindInSolutionExplorerCommandSet, FindInSolutionExplorerConstants.cmdidFindInSolutionExplorer);
            EventHandler invokeHandler = HandleInvokeFindInSolutionExplorer;
            EventHandler changeHandler = HandleChangeFindInSolutionExplorer;
            EventHandler beforeQueryStatus = HandleBeforeQueryStatusFindInSolutionExplorer;
            _command = new OleMenuCommand(invokeHandler, changeHandler, beforeQueryStatus, id);
            mcs.AddCommand(_command);
        }

        public static EnvDTE80.Window2 FindWindow(EnvDTE80.Windows2 windows, EnvDTE.vsWindowType vsWindowType)
        {
            return windows.Cast<EnvDTE80.Window2>().FirstOrDefault(w => w.Type == vsWindowType);
        }

        private void HandleInvokeFindInSolutionExplorer(object sender, EventArgs e)
        {
            try
            {
                EnvDTE.Property track = ApplicationObject.get_Properties("Environment", "ProjectsAndSolution").Item("TrackFileSelectionInExplorer");
                if (track.Value is bool && !((bool)track.Value))
                {
                    track.Value = true;
                    track.Value = false;
                }

                // Find the Solution Explorer object
                EnvDTE80.Windows2 windows = ApplicationObject.Windows as EnvDTE80.Windows2;
                EnvDTE80.Window2 solutionExplorer = FindWindow(windows, EnvDTE.vsWindowType.vsWindowTypeSolutionExplorer);
                if (solutionExplorer != null)
                    solutionExplorer.Activate();
            }
            catch (Exception ex)
            {
                if (ErrorHandler.IsCriticalException(ex))
                    throw;
            }
        }

        private void HandleChangeFindInSolutionExplorer(object sender, EventArgs e)
        {
        }

        private void HandleBeforeQueryStatusFindInSolutionExplorer(object sender, EventArgs e)
        {
            try
            {
                EnvDTE.Document doc = ApplicationObject.ActiveDocument;

                _command.Supported = true;

                bool enabled = false;
                EnvDTE.ProjectItem projectItem = doc != null ? doc.ProjectItem : null;
                if (projectItem != null)
                {
                    if (projectItem.Document != null)
                    {
                        // normal project documents
                        enabled = true;
                    }
                    else if (projectItem.ContainingProject != null)
                    {
                        // this applies to files in the "Solution Files" folder
                        enabled = projectItem.ContainingProject.Object != null;
                    }
                }

                _command.Enabled = enabled;
            }
            catch (ArgumentException)
            {
                // stupid thing throws if the active window is a C# project properties pane
                _command.Supported = false;
                _command.Enabled = false;
            }
            catch (Exception ex)
            {
                if (ErrorHandler.IsCriticalException(ex))
                    throw;

                _command.Supported = false;
                _command.Enabled = false;
            }
        }
    }
}
