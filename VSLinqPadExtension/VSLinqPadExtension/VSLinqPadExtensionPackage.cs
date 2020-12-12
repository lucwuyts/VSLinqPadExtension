using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Events;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace VSLinqPadExtension
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
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(VSLinqPadExtensionPackage.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists , PackageAutoLoadFlags.BackgroundLoad )]
    public sealed class VSLinqPadExtensionPackage : AsyncPackage
    {
        /// <summary>
        /// VSLinqPadExtensionPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "f55338c0-fab2-4ecc-9389-75ea1088233d";


        /// <summary>
        /// Used to monitor file actions in the LinqProject
        /// </summary>
        FileSystemWatcher Watcher;

        FolderService FolderService;


        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            SolutionEvents.OnAfterOpenSolution += SolutionEvents_OnAfterOpenSolution;

            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);            
            await StartLinqPad.InitializeAsync(this);
        }


        /// <summary>
        /// When a linqpad project exists, add new items in the LinqPad folders to the project.
        /// Add a FileSystemWatcher to monitor file actions in the LinqPad project
        /// </summary>
        private void SolutionEvents_OnAfterOpenSolution(object sender, OpenSolutionEventArgs e)
        {            
            var dte = (EnvDTE80.DTE2)this.GetService(typeof(EnvDTE.DTE));
            FolderService = new FolderService(dte);

            ThreadHelper.ThrowIfNotOnUIThread();
            var LinqPadProject = FolderService.GetLINQPadProject();
            if (LinqPadProject != null)
            {
                Watcher = new FileSystemWatcher();
                Watcher.Path = FolderService.GetPath();
                Watcher.Created += (s, a) => {
                    JoinableTaskFactory.RunAsync(async delegate
                   {
                       await Watcher_CreatedAsync(s, a);
                   });
                };
                Watcher.Deleted += Watcher_Deleted;
                Watcher.IncludeSubdirectories = true;
                Watcher.EnableRaisingEvents = true;

                // add existing files
                FolderService.SolutionAddItems();
            }
        }

        private void Watcher_Deleted(object sender, FileSystemEventArgs e)
        {
            
        }

        private async Task Watcher_CreatedAsync(object sender, FileSystemEventArgs e)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync();
            //ThreadHelper.ThrowIfNotOnUIThread();
            var fileName = e.FullPath;
            var ext = Path.GetExtension(fileName);
            FolderService.AddFile(fileName);
        }
    }
}
