﻿//------------------------------------------------------------------------------
// <copyright file="VSPackage1.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System.Windows.Forms;
using EnvDTE;

namespace VSIXProject2
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
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(VSPackage1.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class VSPackage1 : Package
    {
        /// <summary>
        /// VSPackage1 GUID string.
        /// </summary>
        public const string PackageGuidString = "3b76e61c-fe19-493b-b881-23b5e4fc30a5";


        private DTEEvents b;

        /// <summary>
        /// Initializes a new instance of the <see cref="VSPackage1"/> class.
        /// </summary>
        public VSPackage1()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            try
            {
                MessageBox.Show("VSPackage1.Initialize");
                InitCmd();
                //Command1.Initialize(this);

                //  Application.Current.MainWindow.Loaded 

                DTE dTE = (DTE)this.GetService(typeof(DTE));
                this.b = dTE.Events.DTEEvents;
                //     this.b.OnBeginShutdown += new _dispDTEEvents_OnBeginShutdownEventHandler(this, (UIntPtr)ldftn(a));

                Connector c = new Connector();

                GetGrid();


            }
            catch
            {
                throw;
            }

        }

        private void GetGrid()
        {
            var ms = this.GetService(typeof(IVsMonitorSelection)) as IVsMonitorSelection;

            object obj = null;

            ms.GetCurrentElementValue((int)VSConstants.VSSELELEMID.SEID_WindowFrame, out obj);
            var vf = obj as IVsWindowFrame;


            vf.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out obj);
            vf.GetProperty((int)__VSFPROPID.VSFPROPID_DocData, out obj);
            var control = obj as Control;
        }

        private void InitCmd()
        {
            OleMenuCommandService commandService = this.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(new Guid("e5c36492-00e2-4a31-88c9-ac433af1adc4"), 0x0100);
                var menuItem = new MenuCommand((object sender, EventArgs e) =>
                {
                    GetGrid();

                }, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }


        #endregion
    }
}
