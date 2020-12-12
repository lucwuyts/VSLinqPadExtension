using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSLinqPadExtension
{
    public class FolderService
    {

        DTE2 dte;

        public FolderService(DTE2 dte)
        {
            this.dte = dte;
        }


        public string GetPath()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return Path.GetDirectoryName(dte.Solution.FullName);
        }

        public string GetPath(ProjectItem pItem)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!PropertyExists(pItem.Properties, "FullPath"))
            {
                return "";
            }
            return pItem.Properties.Item("FullPath").Value.ToString();
        }

        public string GetPath(Project pItem)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!PropertyExists(pItem.Properties, "FullPath"))
            {
                return "";
            }
            return pItem.Properties.Item("FullPath").Value.ToString();
        }


        public Project GetLINQPadProject()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Solution2 solution = (Solution2)dte.Solution;
            Project solutionFolder = (from Project item in solution.Projects
                                      where item.Name == Constants.LINQPad
                                      select item
                                      ).FirstOrDefault();
            return solutionFolder;
        }

        public ProjectItem GetSolutionFolder(Project project, string fullPath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // a projectItem has a backslash at the ending of the path
            fullPath += @"\";
            foreach (ProjectItem item in project.ProjectItems)
            {
                var path = GetPath(item);
                if (path.Equals(fullPath))
                {
                    return item;
                }
            }
            return null;
        }


        /// <summary>
        /// Add all LINQ files and subfolders to the LinqPad folder in the solution.
        /// </summary>
        public void SolutionAddItems()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var LinqPadProject = GetLINQPadProject();

            if (LinqPadProject == null)
                throw new Exception($"SolutionFolder {Constants.LINQPad}  not found.");

            var filePath = GetPath(LinqPadProject);
            AddItems(filePath);
        }

        private void AddItems(string filePath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var project = GetLINQPadProject();

            // get all items
            // Only 1 subdirectory level !
            var di = new DirectoryInfo(filePath);

            foreach (var item in di.GetFileSystemInfos())
            {
                if (item is FileInfo)
                {
                    var ext = Path.GetExtension(item.FullName);
                    if (".xml.linq".Contains(ext))
                    {
                        project.ProjectItems.AddFromFile(item.FullName);
                    }
                }
                else if (item is DirectoryInfo)
                {
                    var todo = (from s in "drivers,plugins,queries,snippets".Split(',')
                                where item.FullName.ToLower().Contains(s)
                                select s
                                ).Any();
                    if (todo)
                    {
                        var folder = GetSolutionFolder(project, item.FullName);
                        if (folder != null)
                        {
                            AddFiles(folder, item.FullName);
                        }
                    }
                }
            }
        }





        void AddFiles(ProjectItem folder, string path)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var di = new DirectoryInfo(path);
            // only files
            foreach (var item in di.GetFileSystemInfos())
            {
                if (item is FileInfo)
                {
                    var ext = Path.GetExtension(item.FullName);
                    if (".xml.linq".Contains(ext))
                    {
                        folder.ProjectItems.AddFromFile(item.FullName);
                    }
                }
            }
        }



        bool PropertyExists(Properties properties, string name)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (properties == null)
                return false;
            return (from Property p in properties
                    where p.Name == name
                    select p
                    ).Any();
        }

    }
}
