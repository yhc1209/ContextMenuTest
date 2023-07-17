using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.Runtime.InteropServices;

[assembly:AssemblyDescription("This is a test.")]
[assembly:AssemblyTitle("context menu 2023")]
[assembly:AssemblyCompany("How")]
[assembly:AssemblyKeyFile("mykey.snk")]
[assembly:AssemblyInformationalVersion("This is a testing context menu.")]

namespace testDF
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.DirectoryBackground)]
    [COMServerAssociation(AssociationType.AllFilesAndFolders)]
    public class ContextMenu2023 : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            Debug.WriteLine($"[cmtest] selected path:\n  - {String.Join("\n  - ", SelectedItemPaths)}");
            ContextMenuStrip cns = new ContextMenuStrip();
            ToolStripMenuItem menuitem = new ToolStripMenuItem("CM2023"); // 

            ToolStripMenuItem dditem1 = new ToolStripMenuItem("在此開啟powershell");
            dditem1.Click += (s, e) => OpenPowerShell(FolderPath);
            menuitem.DropDownItems.Add(dditem1);

            ToolStripMenuItem dditem2 = new ToolStripMenuItem("在此開啟cmd");
            dditem2.Click += (s, e) => OpenCmd(FolderPath);
            menuitem.DropDownItems.Add(dditem2);

            ToolStripMenuItem dditem3 = new ToolStripMenuItem("這是啥？");
            dditem3.Click += (s, e) => ShowInfo(SelectedItemPaths.ElementAt(0));
            menuitem.DropDownItems.Add(dditem3);

            if (String.IsNullOrEmpty(FolderPath))
            {
                dditem1.Enabled = false;
                dditem2.Enabled = false;
            }
            if (SelectedItemPaths.Count<String>() != 1)
            {
                dditem3.Enabled = false;
            }
            cns.Items.Add(menuitem);

            return cns;
        }

        private void OpenPowerShell(string path)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = "powershell";
            psi.WorkingDirectory = path;
            Process.Start(psi);
        }
        private void OpenCmd(string path)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = "cmd";
            psi.WorkingDirectory = path;
            Process.Start(psi);
        }
        private void ShowInfo(string path)
        {
            string type;
            if (Directory.Exists(path))
                type = "folder";
            else
            {
                if (Path.GetExtension(path).ToLower() == ".lnk")
                {
                    string msg = String.Empty;
                    ShellLink link = new ShellLink(path);
                    msg += $"dscp: {link.Descriptions}\n";
                    msg += $"target: {link.ExecuteFile}\n";
                    msg += $"args: {link.ExecuteArguments}\n";

                    type = $"lnk file\n------ lnk info. ------\n{msg}";
                }
                else
                    type = $"{Path.GetExtension(path)} file";
            }

            MessageBox.Show(
                $"Path: {path}\nType: {type}",
                path, MessageBoxButtons.OK, MessageBoxIcon.Information
            );
        }
    }
}
