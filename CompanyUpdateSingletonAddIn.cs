using RightNow.AddIns.AddInViews;
using System;
using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CompanyUpdateSingleton
{
    public class CompanyUpdateSingletonAddIn : Panel, IWorkspaceComponent2
    {
        /// <summary>
        /// The current workspace record context.
        /// The system global context.
        /// The workspace design mode flag used for safety checking.
        /// </summary>
        private IRecordContext _recordContext;
        public IGlobalContext _globalContext;
        public bool _inDesignMode;
        private bool _attemptedSaveAndClose = false;
        private bool _attemptedSave = false;

        private string CompanyID = string.Empty;

        private int _CompanyID_CustomField_ID;

        private CompanyUpdateSingletonInterface _CompanyUpdateSingletonAddIn;

        /// <summary>
        /// Default constructor.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        /// <param name="GlobalContext">The current system global context, offers access to environment and user account functions and parameters.</param>
        public CompanyUpdateSingletonAddIn(bool inDesignMode, IRecordContext RecordContext, IGlobalContext GlobalContext, int CompanyID_CustomField_ID)
        {
            //You can't access the IRecordContext when in design mode. This Add-In is build specifically to work with the Contact workspace.
            if (!inDesignMode && RecordContext != null && GlobalContext != null && RecordContext.WorkspaceType == RightNow.AddIns.Common.WorkspaceRecordType.Contact)
            {
                _recordContext = RecordContext;
                _globalContext = GlobalContext;
                _inDesignMode = inDesignMode;
                _CompanyID_CustomField_ID = CompanyID_CustomField_ID;

                _recordContext.Closing += _recordContext_Closing;
                _recordContext.DataLoaded += _recordContext_DataLoaded;
                _recordContext.Saving += _recordContext_Saving;
                _CompanyUpdateSingletonAddIn = new CompanyUpdateSingletonInterface();
            }
            if (inDesignMode)
            {
                _CompanyUpdateSingletonAddIn = new CompanyUpdateSingletonInterface();
            }
            // Outside of safety checked constructor so we are able to configure the add-in control layout in the workspace designer.
        }

        private void _recordContext_Saving(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _attemptedSave = true;
        }

        private void _recordContext_DataLoaded(object sender, EventArgs e)
        {
            if (_recordContext != null)
            {
                //string CompanyID = string.Empty;
                var workspace = (_recordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Contact) as IContact);
                if (workspace.CustomField != null && workspace.CustomField.Any(z => z.CfId == _CompanyID_CustomField_ID) && !String.IsNullOrEmpty(workspace.CustomField.First(z => z.CfId == _CompanyID_CustomField_ID).ValStr))
                {
                    CompanyID = workspace.CustomField.First(z => z.CfId == _CompanyID_CustomField_ID).ValStr;
                }
                if (!string.IsNullOrWhiteSpace(CompanyID))
                {
                    if (_attemptedSaveAndClose)
                    {
                        bool _error = false;
                        switch (UpdateStatusHandler.Instance.GetUpdateStatus(CompanyID).updateCustomerStatus)
                        {
                            case UpdateStatuses.UpdateStatus.Called:
                                _error = true;
                                break;
                            case UpdateStatuses.UpdateStatus.Failure:
                                _error = true;
                                break;
                        }
                        switch (UpdateStatusHandler.Instance.GetUpdateStatus(CompanyID).updateOrgStatus)
                        {
                            case UpdateStatuses.UpdateStatus.Called:
                                _error = true;
                                break;
                            case UpdateStatuses.UpdateStatus.Failure:
                                _error = true;
                                break;
                        }
                        switch (UpdateStatusHandler.Instance.GetUpdateStatus(CompanyID).updateSalesStatus)
                        {
                            case UpdateStatuses.UpdateStatus.Called:
                                _error = true;
                                break;
                            case UpdateStatuses.UpdateStatus.Failure:
                                _error = true;
                                break;
                        }
                        if (!_error)
                        {
                            _recordContext.ExecuteEditorCommand(RightNow.AddIns.Common.EditorCommand.Close);
                        }
                    }
                }
            }
        }

        private void _recordContext_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(CompanyID))
            {
                bool _error = false;
                _attemptedSaveAndClose = false;
                string _errorMsg = "Are you sure you want to close?\n\n\n\n";
                string _busyMsg = " request is still waiting for a response from the server...\n";
                string _failMsg = " failed whilst trying to update...\n";

                switch (UpdateStatusHandler.Instance.GetUpdateStatus(CompanyID).updateCustomerStatus)
                {
                    case UpdateStatuses.UpdateStatus.Called:
                        _error = true;
                        _errorMsg += "Update Customer" + _busyMsg;
                        break;
                    case UpdateStatuses.UpdateStatus.Failure:
                        _error = true;
                        _errorMsg += "Update Customer" + _failMsg;
                        break;
                }
                switch (UpdateStatusHandler.Instance.GetUpdateStatus(CompanyID).updateOrgStatus)
                {
                    case UpdateStatuses.UpdateStatus.Called:
                        _error = true;
                        _errorMsg += "Update Exams" + _busyMsg;
                        break;
                    case UpdateStatuses.UpdateStatus.Failure:
                        _error = true;
                        _errorMsg += "Update Exams" + _failMsg;
                        break;
                }
                switch (UpdateStatusHandler.Instance.GetUpdateStatus(CompanyID).updateSalesStatus)
                {
                    case UpdateStatuses.UpdateStatus.Called:
                        _error = true;
                        _errorMsg += "Update Sales Books" + _busyMsg;
                        break;
                    case UpdateStatuses.UpdateStatus.Failure:
                        _error = true;
                        _errorMsg += "Update Sales Books" + _failMsg;
                        break;
                }
                if (!_attemptedSave)
                {
                    if (_error)
                    {
                        DialogResult _diaErr = MessageBox.Show(_errorMsg, "Update Error", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (_diaErr == DialogResult.Yes)
                        {
                            e.Cancel = false;
                        }
                        else if (_diaErr == DialogResult.No)
                        {
                            e.Cancel = true;
                        }
                    }
                }
                if (_attemptedSave)
                {
                    _attemptedSave = false;
                    _attemptedSaveAndClose = true;
                    if (_error)
                    {
                        e.Cancel = true;
                    }
                }
            }
        }

        /// <summary>
        /// Method called by the Add-In framework to retrieve the control.
        /// </summary>
        /// <returns>The control, typically 'this'.</returns>
        public Control GetControl()
        {
            return this._CompanyUpdateSingletonAddIn;
        }

        /// <summary>
        /// Sets the ReadOnly property of this control.
        /// </summary>
        public bool ReadOnly { get; set; }

        /// <summary>
        /// Method which is called when any Workspace Rule Action is invoked.
        /// </summary>
        /// <param name="ActionName">The name of the Workspace Rule Action that was invoked.</param>
        public void RuleActionInvoked(string ActionName)
        {
        }

        /// <summary>
        /// Method which is called when any Workspace Rule Condition is invoked.
        /// </summary>
        /// <param name="ConditionName">The name of the Workspace Rule Condition that was invoked.</param>
        /// <returns>The result of the condition.</returns>
        public string RuleConditionInvoked(string ConditionName)
        {
            return string.Empty;
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && (_recordContext != null))
            {
            }
            base.Dispose(disposing);
        }
    }

    [AddIn("CompanyUpdateSingleton", Version = "1.0.0.0")]
    public class WorkspaceAddInFactory : IWorkspaceComponentFactory2
    {
        #region IWorkspaceComponentFactory2 Members
        IGlobalContext _globalContext;
        /// <summary>
        /// Method which is invoked by the AddIn framework when the control is created.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        /// <returns>The control which implements the IWorkspaceComponent2 interface.</returns>
        public IWorkspaceComponent2 CreateControl(bool inDesignMode, IRecordContext RecordContext)
        {
            //AddInAttribute a = new AddInAttribute("Test");
            return new CompanyUpdateSingletonAddIn(inDesignMode, RecordContext, _globalContext, CompanyID_CustomField_ID);
        }
        #endregion

        #region IFactoryBase Members
        private int _CompanyID_CustomField_ID;
        [ServerConfigProperty(DefaultValue = "50")]
        public int CompanyID_CustomField_ID
        {
            get
            {
                return _CompanyID_CustomField_ID;
            }
            set
            {
                _CompanyID_CustomField_ID = value;
            }
        }

        /// <summary>
        /// The 16x16 pixel icon to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public Image Image16
        {
            get { return Properties.Resources.kek; }
        }

        /// <summary>
        /// The text to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Text
        {
            get { return "CompanyUpdateSingleton"; }
        }

        /// <summary>
        /// The tooltip displayed when hovering over the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Tooltip
        {
            get { return "CompanyUpdateSingleton Tooltip"; }
        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            _globalContext = GlobalContext;
            return true;
        }

        #endregion
    }
}
