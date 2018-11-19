namespace Addin
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.IO;
    using System.ComponentModel.Composition;
    using System.Drawing;
    using Microsoft.Dynamics.Framework.Tools.Extensibility;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Core;
    using Microsoft.Dynamics.Framework.Tools.Configuration;

    using Metadata = Microsoft.Dynamics.AX.Metadata;
    using System.Diagnostics;


    /// <summary>
    /// TODO: Say a few words about what your AddIn is going to do
    /// </summary>
    [Export(typeof(IMainMenu))]
    public class MainMenuAddIn : MainMenuBase
    {
        #region Member variables
        private const string addinName = "Addin";
        private Metadata.Providers.IMetadataProvider metadataProvider = null;
        private Metadata.Service.IMetaModelService metaModelService = null;
        private IDevelopmentConfiguration configuration = null;
        #endregion

        #region Properties
        /// <summary>
        /// Caption for the menu item. This is what users would see in the menu.
        /// </summary>
        public override string Caption
        {
            get
            {
                return AddinResources.MainMenuAddInCaption;
            }
        }

        /// <summary>
        /// Unique name of the add-in
        /// </summary>
        public override string Name
        {
            get
            {
                return MainMenuAddIn.addinName;
            }
        }

        public Metadata.Providers.IMetadataProvider MetadataProvider
        {
            get
            {
                if (this.metadataProvider == null)
                {
                    this.metadataProvider = DesignMetaModelService.Instance.CurrentMetadataProvider;
                }
                return this.metadataProvider;
            }
        }

        public Metadata.Service.IMetaModelService MetaModelService
        {
            get
            {
                if (this.metaModelService == null)
                {
                    this.metaModelService = DesignMetaModelService.Instance.CurrentMetaModelService;
                }
                return this.metaModelService;
            }
        }

        public IDevelopmentConfiguration Configuration
        {
            get
            {
                if (this.configuration == null)
                {
                    this.configuration = DesignMetaModelService.Instance.CurrentConfiguration;
                }
                return this.configuration;
            }
        }
        #endregion

        #region Callbacks
        /// <summary>
        /// Called when user clicks on the add-in menu
        /// </summary>
        /// <param name="e">The context of the VS tools and metadata</param>
        public override void OnClick(AddinEventArgs e)
        {
            try
            {
                // TODO: Do your magic for your add-in
                DirectoryInfo dirInfo = Directory.GetParent(this.Configuration.SSRSReportingLocation);
                 
                string logPath = Path.Combine(dirInfo.FullName,"LogFiles");

                if (Directory.Exists(logPath))
                {
                    DirectoryInfo dirInfoLogs = new DirectoryInfo(logPath);

                    var latestFile = dirInfoLogs.GetFiles()
                                                .OrderByDescending(f => f.LastWriteTime)
                                                .Where(f => f.Name.StartsWith("ReportServerService__"))
                                                .First();
                    if (latestFile != null)
                    {
                        Process.Start(latestFile.FullName);
                    }
                }
            }
            catch (Exception ex)
            {
                CoreUtility.HandleExceptionWithErrorMessage(ex);
            }
        }
        #endregion
    }
}
