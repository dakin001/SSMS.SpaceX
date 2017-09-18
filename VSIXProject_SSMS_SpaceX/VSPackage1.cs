//------------------------------------------------------------------------------
// <copyright file="VSPackage1.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------
extern alias SSMSSpaceX_Ssms2014;

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
using System.IO;

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
    // 插件加载时间
    //  [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    public sealed class VSPackage1 : Package // Connect:Package
    {
        /// <summary>
        /// VSPackage1 GUID string.
        /// </summary>
        public const string PackageGuidString = "3b76e61c-fe19-493b-b881-23b5e4fc30a5";

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
               // MessageBox.Show("SSMS.SpaceX VSPackage1.Initialize");
                FireRocket();
            }
            catch (Exception ex)
            {
                MessageBox.Show("VSPackage1.Initialize exceptio:" + ex.ToString());
            }

            DelayAddSkipLoadingReg();
        }

        private object FireRocket()
        {
            var ssmsInterfacesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SqlWorkbench.Interfaces.dll");

            if (File.Exists(ssmsInterfacesPath))
            {
                var ssmsInterfacesVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(ssmsInterfacesPath);

                switch (ssmsInterfacesVersion.FileMajorPart)
                {
                    //case 14:
                    //    return new Ssms2017::SSMSSpaceX.Command1.Initialize(this);;
                    case 13:
                        return SSMSSpaceX.Command1.Initialize(this); ;
                    case 12:
                        return SSMSSpaceX_Ssms2014::SSMSSpaceX.Command1.Initialize(this); ;
                    default:
                        break;
                }
            }

            return SSMSSpaceX.Command1.Initialize(this);

        }


        private void DelayAddSkipLoadingReg()
        {
            var delay = new Timer();
            delay.Tick += delegate (object o, EventArgs e)
            {
                delay.Stop();
                AddSkipLoadingReg();
            };
            delay.Interval = 1000;
            delay.Start();
        }

        private void AddSkipLoadingReg()
        {
            // Reg setting is removed after initialize. Wait short delay then recreate it.
            var myPackage = this.UserRegistryRoot.CreateSubKey(@"Packages\{" + PackageGuidString + "}");
            myPackage.SetValue("SkipLoading", 1);
        }

        #endregion
    }
}
