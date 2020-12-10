using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using Task = System.Threading.Tasks.Task;

namespace VSLinqPadExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class StartLinqPad
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a15da661-356a-4878-a9e8-36c37a7f38d5");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        static IProjectFolderService projectFolderService;

        static DTE2 dte;

        /// <summary>
        /// Initializes a new instance of the <see cref="StartLinqPad"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private StartLinqPad(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);

        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static StartLinqPad Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in StartLinqPad's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new StartLinqPad(package, commandService);

            projectFolderService = await package.GetServiceAsync(typeof(IProjectFolderService)) as IProjectFolderService;


        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            dte = (DTE2)(this.ServiceProvider.GetServiceAsync(typeof(DTE)).GetAwaiter().GetResult());
            try
            {
                var solutionDir = System.IO.Path.GetDirectoryName(dte.Solution.FullName);

                // Ensure linqpad project dir and folders exist.
                var linqPad = Path.Combine(solutionDir, "LinqPad");
                var linqPad_Drivers = Path.Combine(linqPad, "drivers");
                var linqPad_Plugins = Path.Combine(linqPad, "plugins");
                var linqPad_Queries = Path.Combine(linqPad, "Queries");
                var linqPad_Snippets = Path.Combine(linqPad, "snippets");
                Directory.CreateDirectory(linqPad);
                Directory.CreateDirectory(linqPad_Drivers);
                Directory.CreateDirectory(linqPad_Plugins);
                Directory.CreateDirectory(linqPad_Queries);
                Directory.CreateDirectory(linqPad_Snippets);


                EnsureFileExists(linqPad, "LINQPad6.exe");
                EnsureFileExists(linqPad, "LINQPad.GUI.dll");
                EnsureFileExists(linqPad, "LINQPad.GUI.runtimeconfig.json");
                EnsureFileExists(linqPad, "LINQPad.Runtime.dll");
                EnsureFileExists(linqPad, "LINQPad.Runtime.runtimeconfig.json");

                MakeConnectionsXml(linqPad);

                MakeSolutionFolderWithItems(dte, linqPad);

                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = Path.Combine(linqPad, "LINQPad6.exe");
                proc.Start();
            }
            catch (Exception ex)
            {
                dte.StatusBar.Text = $"Error: {ex.Message}";
            }
        }

        private void MakeSolutionFolderWithItems(DTE2 dte, string linqPadFolder)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var folderName = "LINQPad";

            Solution2 solution = (Solution2)dte.Solution;

            Project solutionFolder = null;

            foreach (Project item in solution.Projects)
            {
                if (item.Name == folderName)
                {
                    solutionFolder = item;
                    break;
                }
            }
            // create if not exist
            if (solutionFolder == null)
            {
                solutionFolder = solution.AddSolutionFolder("LINQPad");
                IncludeFiles(linqPadFolder, solutionFolder);
            }
            
        }

        private void IncludeFiles(string folder, Project project)
        {
            var validExtentions = new string[] { "xml", "linq" };

            ThreadHelper.ThrowIfNotOnUIThread();
            var di = new DirectoryInfo(folder);
            foreach (var item in di.GetFileSystemInfos())
            {
                if (item is FileInfo)
                {
                    // only LINQ files
                    if (validExtentions.Any(s => s.IndexOf(item.Name, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        project.ProjectItems.AddFromFile(item.FullName);
                    }
                }
                else if (item is DirectoryInfo)
                {
                    SolutionFolder solutionFolder = (SolutionFolder)project.Object;
                    var newSolutionFolder = solutionFolder.AddSolutionFolder(item.Name);
                    IncludeFiles(item.FullName, newSolutionFolder);
                }
            }
        }

        private void MakeConnectionsXml(string pathToDest)
        {
            var dest = Path.Combine(pathToDest, "ConnectionsV2.xml");
            if( !File.Exists(dest))
            {
                File.WriteAllText(dest, "<?xml version=\"1.0\" encoding=\"utf-8\"?><Connections></Connections>");
            }
        }

        void EnsureFileExists(string pathToDest, string filename)
        {
            var srcLinqPad = @"C:\Program Files\LINQPad6";
            var srcFile = Path.Combine(srcLinqPad, filename );
            if (!File.Exists(srcFile))
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "LINQPad6 not found.",
                    $"Path: {srcFile}",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                throw new Exception("LINQPad6 not found.");
            }
            else
            {
                var destFile = Path.Combine(pathToDest, filename);
                if(!File.Exists(destFile))
                {
                    File.Copy(srcFile, destFile);
                }                
            }
        }

    }
}
