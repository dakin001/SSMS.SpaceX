//------------------------------------------------------------------------------
// <copyright file="Command1.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using System.Windows.Forms;
using Microsoft.SqlServer.Management.UI.Grid;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;

namespace SSMSSpaceX
{
    /// <summary>
    /// Command handler
    /// </summary>
    public sealed class Command1
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("e5c36492-00e2-4a31-88c9-ac433af1adc4");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private Command1(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            //   CreateMenu();
            CreateGridMenu();
        }

        private void CreateMenu()
        {
            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command1 Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
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
        public static Command1 Initialize(Package package)
        {
            Instance = new Command1(package);
            return Instance;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "Command Test 1";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private DTE2 DTE
        {
            get { return (DTE2)this.ServiceProvider.GetService(typeof(DTE)); }
        }

        private void CreateGridMenu()
        {
            CommandBar tabContext = ((CommandBars)this.DTE.CommandBars)["SQL Results Grid Tab Context"];
            //CommandBarPopup menuSpaceX = (CommandBarPopup)tabContext.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            //menuSpaceX.Caption = "SpaceX";
            //menuSpaceX.BeginGroup = true;
            var btnSaveToScript = tabContext.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true) as CommandBarButton;
            btnSaveToScript.Caption = "SpaceX: Save data to script";
            btnSaveToScript.Click += (CommandBarButton Ctrl, ref bool CancelDefault) =>
            {
                ScriptGrid(SetDocSql);
            };

            var btnSaveToExcel = tabContext.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true) as CommandBarButton;
            btnSaveToExcel.Caption = "SpaceX: Copy data script";
            btnSaveToExcel.Click += (CommandBarButton Ctrl, ref bool CancelDefault) =>
            {
                ScriptGrid(sql =>
                {
                    Clipboard.SetDataObject(sql);
                });
            };
        }


        private void ScriptGrid(Action<string> action = null)
        {
            var ms = this.ServiceProvider.GetService(typeof(IVsMonitorSelection)) as IVsMonitorSelection;

            object obj = null;

            ms.GetCurrentElementValue((int)VSConstants.VSSELELEMID.SEID_WindowFrame, out obj);
            var vf = obj as IVsWindowFrame;

            vf.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out obj);
            var control = obj as Control;
            if (control != null)
            {
                //   Microsoft.SqlServer.Management.UI.VSIntegration.Editors.SqlScriptEditorControl;
                var gridControl = (GridControl)((ContainerControl)((ContainerControl)control).ActiveControl).ActiveControl;
                GetGridScriptData(gridControl, action);
            }
        }

        private void GetGridScriptData(GridControl gridControl, Action<string> action)
        {
            if (gridControl == null)
            {
                return;
            }
            long totalRows = gridControl.GridStorage.NumRows();
            if (totalRows <= 5000L || MessageBox.Show(string.Concat(new object[]
            {
                            "Result set number ",
                            " contains a lot of rows. Scripting it might cause Out Of Memory exceptions.",
                            Environment.NewLine,
                            "Do you want to try to script it?"
            }), "Script Grid Results", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) != DialogResult.No)
            {
                List<string> columnHeaderList;
                List<List<string>> data;
                GetGridData(gridControl, out columnHeaderList, out data);



                string result = string.Join("\r\n,", data.Select(line => $"({string.Join(", ", line)})"));

                result = $"-- INSERT INTO #tmp_GridResults ({string.Join(", ", columnHeaderList)})\r\n"
                    + $"select * from(values \r\n {result}\r\n) as T({string.Join(", ", columnHeaderList)})";

                action(result);
            }
        }

        private static void GetGridData(GridControl gridControl, out List<string> columnHeaderList, out List<List<string>> data)
        {
            int columnsNumber = gridControl.ColumnsNumber;
            long totalRows = gridControl.GridStorage.NumRows();

            columnHeaderList = new List<string>(columnsNumber);
            for (int j = 1; j < columnsNumber; j++)
            {
                string text;
                gridControl.GetHeaderInfo(j, out text, out Bitmap bitmap);
                if (columnHeaderList.Contains("[" + text + "]"))
                {
                    text = text + "_" + j.ToString();
                }
                columnHeaderList.Add("[" + text + "]");
            }

            data = new List<List<string>>();

            for (long rowNum = 0L; rowNum < totalRows; rowNum += 1L)
            {
                var row = new List<string>();
                for (int colNum = 1; colNum < columnsNumber; colNum++)
                {
                    string cellText = gridControl.GridStorage.GetCellDataAsString(rowNum, colNum) ?? "";
                    cellText = cellText.Replace("'", "''");
                    if (true)
                    {
                        cellText = cellText.Trim();
                    }
                    if (cellText.Equals("NULL", StringComparison.OrdinalIgnoreCase))
                    {
                        //if (!true)
                        //{
                        //    cellText = "'" + cellText + "'";
                        //}
                    }
                    else
                    {
                        cellText = "N'" + cellText + "'";
                    }
                    row.Add(cellText);
                }

                data.Add(row);
            }
        }

        private string GetExecSQL()
        {
            string contentSql = string.Empty;
            var textDoc = (TextDocument)this.DTE.ActiveDocument.Object();
            contentSql = textDoc.Selection.Text;
            if (string.IsNullOrWhiteSpace(contentSql))
            {
                contentSql = textDoc.StartPoint.CreateEditPoint().GetText(textDoc.EndPoint);
            }

            return contentSql;
        }
        private void SetDocSql(string sql)
        {
            Microsoft.SqlServer.Management.UI.VSIntegration.Editors.ScriptFactory.Instance.CreateNewBlankScript(Microsoft.SqlServer.Management.UI.VSIntegration.Editors.ScriptType.Sql);
            var textDoc = (TextDocument)this.DTE.ActiveDocument.Object(null);
            textDoc.EndPoint.CreateEditPoint().Insert(sql);
        }
    }
}
