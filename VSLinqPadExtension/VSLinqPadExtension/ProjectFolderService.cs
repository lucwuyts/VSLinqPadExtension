using EnvDTE80;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSLinqPadExtension
{

    /// <summary>
    /// copied from: https://github.com/conwid/VSCleanBin/blob/master/CleanBinCommands/Services/ProjectFolderSerivce.cs
    /// </summary>
    public class ProjectFolderService : IProjectFolderService
    {

        DTE2 dte;

        public ProjectFolderService(DTE2 dte)
        {
            this.dte = dte;
        }

        public string GetSolutionRootFolder()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrEmpty(dte.Solution.FullName))
                return null;
            return File.Exists(dte.Solution.FullName) ? Path.GetDirectoryName(dte.Solution.FullName) : null;

        }


    }
}
