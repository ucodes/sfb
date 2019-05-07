using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Management.Automation.Runspaces;
using System.Management.Automation.Remoting;
using System.Security;
using System.Threading;
using System.Windows.Threading;
using System.Collections;
using System.Text.RegularExpressions;

namespace SfBtool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //declare runspace for powershell
        static Runspace runspace;

        //timer to prevent PS session timeout
        System.Timers.Timer Idletimer;
        int PSSessionTimeOut = 45 * 60 * 1000; //in ms
        short NumberOfPSReconnectionsCurrent = 0, NumberOfPSReconnectionsThreshold = 14;

        static bool ConnectionFlag = false;

        //user attributes
        string Identity, SipAddress, RegistrarPool, LineURI, ADPhone;

        Hashtable UserPoliciesOfFoundUser = new Hashtable();
        bool EnterpriseVoiceEnabled, HostedVoicemail, AllowInternational;

        //tenant policies
        List<string> TenantDialPlans = new List<string>();
        List<string> TenantConferencingPolicies = new List<string>();
        List<string> TenantVoicePolicies = new List<string>();
        List<string> TenantExternalAccessPolicies = new List<string>();
        List<string> TenantHostedVoicemailPolicies = new List<string>();
        List<string> TenantMobilityPolicies = new List<string>();

        List<string> TenantHybridPSTNSites = new List<string>();

        private string password;
        private string userName;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            // show login window;
            LoginWindow subWindow = new LoginWindow();
            subWindow.Owner = this;
            subWindow.ShowDialog();

            //Ok button clicked on login window
            if (subWindow.Confirmed && !String.IsNullOrEmpty(subWindow.Password) 
                               && !String.IsNullOrEmpty(subWindow.Login))
            {
                password = subWindow.Password;
                userName = subWindow.Login;
                ConnectButton.Content = "Connecting...";
                ConnectButton.IsEnabled = false;
                MainText.AppendText("Connecting to remote PowerShell, wait...\n");

                //create background thread and start PSconnection   
                Thread t1 = new Thread(new ThreadStart(PSconnection));
                t1.Start();
            }
        }

 private void PSconnection()
        {
            //clear tenant policies from previous connection
            TenantDialPlans.Clear();
            TenantHybridPSTNSites.Clear();
            TenantConferencingPolicies.Clear();
            TenantVoicePolicies.Clear();
            TenantExternalAccessPolicies.Clear();
            TenantHostedVoicemailPolicies.Clear();
            TenantMobilityPolicies.Clear();

            string returnstring="";
            ConnectionFlag = false;

         //  System.Uri uri = new Uri("https://localhost:12346/ocspowershell");
            System.Security.SecureString securePassword = String2SecureString(password);

            PSCredential creds = new PSCredential(userName, securePassword);

            runspace = RunspaceFactory.CreateRunspace();

            PowerShell powershell = PowerShell.Create();
            PSCommand command = new PSCommand();
            command.AddCommand("New-CsOnlineSession");
         //  command.AddParameter("ConnectionUri", uri);
            command.AddParameter("Credential", creds);
         //   command.AddParameter("Authentication", "Default");
            PSSessionOption sessionOption = new PSSessionOption();
            sessionOption.SkipCACheck = true;
            sessionOption.SkipCNCheck = true;
            sessionOption.SkipRevocationCheck = true;
            TimeSpan ts = TimeSpan.FromMinutes(5);
            sessionOption.IdleTimeout = ts;
            sessionOption.CancelTimeout = ts;
            sessionOption.OperationTimeout = ts;
            sessionOption.OpenTimeout = TimeSpan.FromMinutes(1);
         //   command.AddParameter("SessionOption", sessionOption);

            powershell.Commands = command;

            try
            {
                // open the runspace (connect to local powershell)
                runspace.Open();

                // associate the runspace with powershell
                powershell.Runspace = runspace;

                // invoke the powershell to obtain the results
                Collection<PSSession> result = powershell.Invoke<PSSession>();

                foreach (ErrorRecord current in powershell.Streams.Error)
                {
                    returnstring += "Creating new PSSession error_start\n";
                    returnstring += "Exception: " + current.Exception.ToString() + "\n";
                    returnstring += "Inner Exception: " + current.Exception.InnerException + "\n";
                    returnstring += "Creating new PSSession error_end\n";
                    ConnectionFlag = false;
                    return;
                }

                //one PS sessions is expected
                if (result.Count != 1)
                {
                    returnstring += "Couldnt connect to SfBO or unexpected number " +
                        "of Remote Runspace connections returned" + "\n";
                    //throw new Exception("Unexpected number of Remote Runspace connections returned.");
                }

                else
                {
                    AppendMainText("New-PSSession created. Importing PS session, wait..." + "\n");
                }

                // Set the runspace as a local variable on the runspace
                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "ra");
                command.AddParameter("Value", result[0]);
                powershell.Commands = command;
                powershell.Runspace = runspace;

                powershell.Invoke();

                foreach (ErrorRecord current in powershell.Streams.Error)
                {
                    returnstring += "Set-Variable error_start\n";
                    returnstring += "Exception: " + current.Exception.ToString() + "\n";
                    returnstring += "Inner Exception: " + current.Exception.InnerException + "\n";
                    returnstring += "Set-Variable error_end\n";
                    ConnectionFlag = false;
                    return;
                }

                // set PS execution policy to unrestricted for the session
                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Set-ExecutionPolicy");
                command.AddParameter("ExecutionPolicy", "Unrestricted");
                command.AddParameter("Scope", "CurrentUser");
                command.AddParameter("Force");
                powershell.Commands = command;
                powershell.Runspace = runspace;

                powershell.Invoke();

                foreach (ErrorRecord current in powershell.Streams.Error)
                {
                    returnstring += "Set-ExecutionPolicy error_start\n";
                    returnstring += "Exception: " + current.Exception.ToString() + "\n";
                    returnstring += "Inner Exception: " + current.Exception.InnerException + "\n";
                    returnstring += "Set-ExecutionPolicy error_end\n";
                    ConnectionFlag = false;
                    return;
                }

                //  import Lync cmdlets from the remote server in the current local runspace (using Import-PSSession)
                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("Import-PSSession -Session $ra");
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();

                foreach (ErrorRecord current in powershell.Streams.Error)
                {
                    returnstring += "Import-PSSession error_start\n";
                    returnstring += "Exception: " + current.Exception.ToString() + "\n";
                    returnstring += "Inner Exception: " + current.Exception.InnerException + "\n";
                    returnstring += "Import-PSSession error_end\n";
                    ConnectionFlag = false;
                    return;
                }

                // Retrieve server info
                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("Get-CsTenant");
                powershell.Commands = command;
                powershell.Runspace = runspace;

                Collection<PSObject> results = new Collection<PSObject>();
                results = powershell.Invoke();

                foreach (ErrorRecord current in powershell.Streams.Error)
                {
                    returnstring += "Get-CsWindowsService error_start\n";
                    returnstring += "Exception: " + current.Exception.ToString() + "\n";
                    returnstring += "Inner Exception: " + current.Exception.InnerException + "\n";
                    returnstring += "Get-CsWindowsService error_end\n";
                    ConnectionFlag = false;
                    return;
                }

                //successfull connection
                if (powershell.Streams.Error.Count == 0)
                {
                    AppendMainText("Successfully connected to " +
                            results.First().Properties["DisplayName"].Value.ToString() + "tenant " + "\n");
                    //retrive tenant policies info 
                    AppendMainText("Getting tenant information" + "\n");
                    GetPolicyInfo();
                    AppendMainText("Search for a user and " +
                                            "change required attributes" + ".\n") ;
                    ConnectionFlag = true;
                    //update buttons
                    UpdateButton("ConnectButton", false, "Connected");
                    UpdateButton("SearchButton", true, "");

                        //start timer to prevent a session timeout
                        Idletimer = new System.Timers.Timer(PSSessionTimeOut);
                        Idletimer.Elapsed += new System.Timers.ElapsedEventHandler(OnElapsed);
                        Idletimer.AutoReset = false;
                        Idletimer.Start();

                }

                //unsuccessfull connection (the error will be returned in finally block)
                else
                {
                    ConnectionFlag = false;
                }

            }

            catch (Exception ex)
            {
                returnstring += "Following exception happened while connecting to remote PS:\n";
                returnstring += "Exception: " + ex.Message.ToString() + "\n";
                ConnectionFlag = false;
            }

            finally
            {

                // Finally dispose the powershell and set all variables to null to free
                // up any resources.
                powershell.Dispose();
                powershell = null;
                AppendMainText(returnstring + "\n");

                if (!ConnectionFlag)
                {
                    //PSExecute("Remove-Variable ra");
                    PSDisconnect();
                    UpdateButton("ConnectButton", true, "Connect");
                }

            }
        }

        //disconnect current PS session and disable buttons
        private void PSDisconnect()
        {
            //disable functional buttons
            UpdateButton("ConnectButton", false, "");
            UpdateButton("SearchButton", false, "");
            UpdateButton("ViewConferencingPolicyButton", false, "");
            UpdateButton("ViewTenantDialPlanButton", false, "");
            UpdateButton("ViewExternalAccessPolicyButton", false, "");
            UpdateButton("ViewHostedVoicemailPolicyButton", false, "");
            UpdateButton("ViewMobilityPolicyButton", false, "");
            UpdateButton("ViewVoicePolicyButton", false, "");
            UpdateButton("UpdateUserButton", false, "");
            UpdateButton("ResetPinButton", false, "");
            UpdateButton("GetPhoneFromAD", false, "");

            ConnectionFlag = false;


            PSExecute("Get-PSSession | Remove-PSSession");
            // dispose the runspace and enable garbage collection
            if (runspace != null)
            {
                runspace.Dispose();
                runspace = null;
            }

            AppendMainText("\nPS session was removed\n");

        }

    private void Idletimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            throw new NotImplementedException();
        }

        //retrive policies info to fill combo boxes items  
        private void GetPolicyInfo()
        {
            Collection<PSObject> results = new Collection<PSObject>();

            try
            {
                results = PSExecute("Get-CsTenantDialPlan");
                foreach (PSObject PSresult in results)
                {
                    TenantDialPlans.Add(PSresult.Properties["Identity"].Value.ToString());
                }               
                results = PSExecute("Get-CsHybridPSTNSite");
                //adding "None" as the first HybridPSTNSite
                TenantHybridPSTNSites.Add("None");
                if (results != null)
                {
                    foreach (PSObject PSresult in results)
                    {
                        TenantHybridPSTNSites.Add(PSresult.Properties["Identity"].Value.ToString());
                    }
                }
                results = PSExecute("Get-CsConferencingPolicy");
                foreach (PSObject PSresult in results)
                {
                    TenantConferencingPolicies.Add(PSresult.Properties["Identity"].Value.ToString());
                }
                results = PSExecute("Get-CsVoicePolicy");
                foreach (PSObject PSresult in results)
                {
                    TenantVoicePolicies.Add(PSresult.Properties["Identity"].Value.ToString());
                }
                results = PSExecute("Get-CsExternalAccessPolicy");
                foreach (PSObject PSresult in results)
                {
                    TenantExternalAccessPolicies.Add(PSresult.Properties["Identity"].Value.ToString());
                }
                results = PSExecute("Get-CsHostedVoicemailPolicy");
                foreach (PSObject PSresult in results)
                {
                    TenantHostedVoicemailPolicies.Add(PSresult.Properties["Identity"].Value.ToString());
                }
                results = PSExecute("Get-CsMobilityPolicy");
                foreach (PSObject PSresult in results)
                {
                    TenantMobilityPolicies.Add(PSresult.Properties["Identity"].Value.ToString());
                }
            }

            catch (Exception ex)
            {
                AppendMainText("Following error happened while retrieving policies:\n");
                AppendMainText("Exception: " + ex.Message.ToString() + "\n");
            }

        }
   
    private Collection<PSObject> PSExecute(string Script)
    {
        try
        {
            PowerShell powershell = PowerShell.Create();
            PSCommand command = new PSCommand();
               
            powershell = PowerShell.Create();
            command = new PSCommand();
            command.AddScript(Script);
            powershell.Commands = command;
            powershell.Runspace = runspace;

            Collection<PSObject> results = new Collection<PSObject>();
            results = powershell.Invoke();

            foreach (ErrorRecord current in powershell.Streams.Error)
                {
                    AppendMainText("\n>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>\nPSExecute error start\n");
                    AppendMainText("Exception: " + current.Exception.ToString() + "\n");
                    AppendMainText("Inner Exception: " + current.Exception.InnerException + "\n");
                    AppendMainText("PSExecute error end\n>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>\n");
                    return null;
                }

                return results;

        }

            catch (Exception ex)
            {
                AppendMainText("Following exception happened when calling PS command:\n");
                AppendMainText("Exception: " + ex.Message.ToString());
                return null;
            }

            finally
        {
                // dispose the runspace and enable garbage collection
                //  runspace.Dispose();
                //  runspace = null;

                // Finally dispose the powershell and set all variables to null to free
                // up any resources.
                // powershell.Dispose();
                // powershell = null;    
         }
    }

private static SecureString String2SecureString(string password)
        {
            SecureString remotePassword = new SecureString();
            for (int i = 0; i < password.Length; i++)
                remotePassword.AppendChar(password[i]);
            return remotePassword;
        }


        //search button click
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {

            //AppendMainText("\nSearching, wait...");

            //clean attributes from previous user
            UserPoliciesOfFoundUser.Clear();
            Identity = "";
            SipAddress = "";
            RegistrarPool = "";
            LineURI = "";
            ADPhone = "";
            EnterpriseVoiceEnabled = false;
            HostedVoicemail = false;
            AllowInternational = false;

            TenantDialPlanComboBox.ItemsSource = null;
            HybridPstnSiteComboBox.ItemsSource = null;
            ConferencingPolicyComboBox.ItemsSource = null;
            VoicePolicyComboBox.ItemsSource = null;
            ExternalAccessPolicyComboBox.ItemsSource = null;
            HostedVoicemailPolicyComboBox.ItemsSource = null;
            MobilityPolicyComboBox.ItemsSource = null;

            //update UI
            UpdateUIwithUserAttributes();

            //disable functional buttons
            UpdateButton("ViewTenantDialPlanButton", false, "");
            UpdateButton("ViewConferencingPolicyButton", false, "");
            UpdateButton("ViewExternalAccessPolicyButton", false, "");
            UpdateButton("ViewHostedVoicemailPolicyButton", false, "");
            UpdateButton("ViewMobilityPolicyButton", false, "");
            UpdateButton("ViewVoicePolicyButton", false, "");
            UpdateButton("UpdateUserButton", false, "");
            UpdateButton("ResetPinButton", false, "");
            UpdateButton("GetPhoneFromAD", false, "");

            Collection<PSObject> results = new Collection<PSObject>();

            //search by OnPremLineURI if choosen
            if (SearchByPhoneCheckBox.IsChecked == true)
            {
                //if wildcard is specified in the begining of the phone string 
                if (SearchUserTextBox.Text.StartsWith("*"))
                {
                    results = PSExecute("Get-CsOnlineUser -filter{OnPremLineURI -like \"" + SearchUserTextBox.Text + "\"}");
                }
                else
                {
                    results = PSExecute("Get-CsOnlineUser -filter{OnPremLineURI -like \"tel:" + SearchUserTextBox.Text + "\"}");
                }
            }
            else
            {
                results = PSExecute("Get-CsOnlineUser \"" + SearchUserTextBox.Text + "\"");
            }

            try
            {
                if (results != null)
                {
                    //check whether more than one user was found
                    if (results.Count > 1)
                    {
                        AppendMainText("\nFollowing users were found, please type exact SIP address or number of a user" + "\n");

                        foreach (PSObject PSresult in results)
                        {
                            AppendMainText(PSresult.Properties["SipAddress"].Value.ToString() + "\n");
                            AppendMainText(PSresult.Properties["OnPremLineURI"].Value.ToString() + "\n");
                        }
                    }
                    //one user is expected
                    else if (results.Count == 1)
                    {


                        AppendMainText("\nFollowing user was found" + "\n");
                        //show only following attributes of a user in main window text box
                        string[] UserAttributesToDisplay = { "Company", "Department", "Description", "IPPhone", "City", "MobilePhone",
                                        "Office", "Title", "Phone", "UsageLocation", "OnPremLineURI", "UserPrincipalName", "DisplayName", "SipAddress",
                                        "Enabled", "EnterpriseVoiceEnabled", "HostedVoicemail", "DialPlan", "TenantDialPlan" };
                        foreach (string AttrName in UserAttributesToDisplay)
                        {
                            AppendMainText(AttrName + " : " + results.First().Properties[AttrName].Value + "\n");
                        }

                        /*
                        //show all user info in the main box
                        AppendMainText("\nAll attributes of found user:" + "\n");
                        foreach (PSPropertyInfo PSpr in results.First().Properties)
                        {
                            AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                        }
                        */

                        //get user info
                        Identity = results.First().Properties["Identity"].Value.ToString();
                        SipAddress = results.First().Properties["SipAddress"].Value.ToString();
                        //if the user sfbo enabled, return pool name, otherwise "The user isn't Enabled" string
                        RegistrarPool = (results.First().Properties["Enabled"].Value.ToString() == "True") ?
                            results.First().Properties["RegistrarPool"].Value.ToString() : "The user isn't Enabled";
                        //get AD Phone attribute, remove spaces if any and add + sign if its absent
                        ADPhone = results.First().Properties["Phone"].Value.ToString();
                        ADPhone = String.IsNullOrEmpty(results.First().Properties["Phone"].Value.ToString()) ?
                                   "" : Regex.Replace(ADPhone, " ", "");
                        if ((!ADPhone.StartsWith("+")) && (ADPhone != ""))
                        {
                            ADPhone = "+" + ADPhone;
                        }
                        string _tempHybridPstnSiteName = "None"; //the value will be assigned later  
                        //if policies are null or empty, the user has Global policy
                        string _tempTenantDialPlan = (results.First().Properties["TenantDialPlan"].Value == null) ?
                                    "Global" : "Tag:" + results.First().Properties["TenantDialPlan"].Value.ToString();
                        string _tempConferencingPolicy = results.First().Properties["ConferencingPolicy"].Value == null ?
                                    "Global" : "Tag:" + results.First().Properties["ConferencingPolicy"].Value.ToString();
                        string _tempVoicePolicy = (results.First().Properties["VoicePolicy"].Value == null) ?
                                    "Global" : "Tag:" + results.First().Properties["VoicePolicy"].Value.ToString();
                        string _tempExternalAccessPolicy = (results.First().Properties["ExternalAccessPolicy"].Value == null) ?
                                    "Global" : "Tag:" + results.First().Properties["ExternalAccessPolicy"].Value.ToString();
                        string _tempHostedVoicemailPolicy = (results.First().Properties["HostedVoicemailPolicy"].Value == null) ? 
                                    "Global" : "Tag:" + results.First().Properties["HostedVoicemailPolicy"].Value.ToString();
                        string _tempMobilityPolicy = (results.First().Properties["MobilityPolicy"].Value == null) ?
                                    "Global" : "Tag:" + results.First().Properties["MobilityPolicy"].Value.ToString();
                        if (results.First().Properties["EnterpriseVoiceEnabled"].Value.ToString() == "True")
                        {
                            EnterpriseVoiceEnabled = true;
                        }
                        //international calls will be allowed if OnlineDialOutPolicy is empty or has DialoutCPCandPSTNInternational value
                        if (results.First().Properties["OnlineDialOutPolicy"].Value == null)
                        {
                            AllowInternational = true;
                        }
                        else
                        {
                            if (results.First().Properties["OnlineDialOutPolicy"].Value.ToString() == "DialoutCPCandPSTNInternational")
                            {
                                AllowInternational = true;
                            }
                        }
                        if (results.First().Properties["HostedVoiceMail"].Value != null)
                        {
                            if (results.First().Properties["HostedVoiceMail"].Value.ToString() == "True")
                            {
                                HostedVoicemail = true;
                            }
                        }


                        //check - if lineuri is null set up an empty string 
                        if (results.First().Properties["OnPremLineURI"].Value == null)
                        {
                            LineURI = "";
                        }
                        else
                        {
                            LineURI = results.First().Properties["OnPremLineURI"].Value.ToString();
                        }

                        //getting user's Hybrid Pstn Site Name
                        results = PSExecute("Get-CsUserPstnSettings -Identity \"" + Identity + "\"");
                        //one object is expected
                        if (results.Count == 1)
                        {
                            //if the user doesnt have a HybridPstnSite assigned, return "None"
                            _tempHybridPstnSiteName = String.IsNullOrEmpty(results.First().Properties["HybridPstnSiteName"].Value.ToString()) ?
                                   _tempHybridPstnSiteName = "None" : _tempHybridPstnSiteName = results.First().Properties["HybridPstnSiteName"].Value.ToString();
                        }

                        //add effective user policies to user's hashtable
                        UserPoliciesOfFoundUser.Add("TenantDialPlan", _tempTenantDialPlan);
                        UserPoliciesOfFoundUser.Add("HybridPstnSiteName", _tempHybridPstnSiteName);
                        UserPoliciesOfFoundUser.Add("ConferencingPolicy", _tempConferencingPolicy);
                        UserPoliciesOfFoundUser.Add("VoicePolicy", _tempVoicePolicy);
                        UserPoliciesOfFoundUser.Add("ExternalAccessPolicy", _tempExternalAccessPolicy);
                        UserPoliciesOfFoundUser.Add("HostedVoicemailPolicy", _tempHostedVoicemailPolicy);
                        UserPoliciesOfFoundUser.Add("MobilityPolicy", _tempMobilityPolicy);

                        //update UI (fill neccessary attributes in action block (left section)) and fill's combo boxes items with tenant policies 
                        UpdateUIwithUserAttributes();
                        //enable functional buttons
                        UpdateButton("ViewTenantDialPlanButton", true, "");
                        UpdateButton("ViewConferencingPolicyButton", true, "");
                        UpdateButton("ViewExternalAccessPolicyButton", true, "");
                        UpdateButton("ViewHostedVoicemailPolicyButton", true, "");
                        UpdateButton("ViewMobilityPolicyButton", true, "");
                        UpdateButton("ViewVoicePolicyButton", true, "");
                        UpdateButton("UpdateUserButton", true, "Update");
                        UpdateButton("ResetPinButton", true, "");
                        UpdateButton("GetPhoneFromAD", true, "");

                    }

                    else
                    {
                        AppendMainText("\nNo users were found\n");
                    }
                }
                else
                {
                    AppendMainText("\nSome error occured, try again\n");
                }
             
            }

            catch (Exception ex)
            {
                AppendMainText("Following exception occured while searching Lync users:\n");
                AppendMainText("Exception: " + ex.Message.ToString());
            }
        }

        private void UpdateUserButton_Click(object sender, RoutedEventArgs e)
        {
            bool IsSipAddressChanged = false,
                 IsHybridPstnSiteNameChanged = false,
                 //   IsRegistrarPoolChanged = false, 
                 IsLineURIChanged = false,
                 IsEnterpriseVoiceEnabledChanged = false,
                 IsHostedVoicemailChanged = false,
                 IsConferencingPolicyChanged = false,
                 IsVoicePolicyChanged = false,
                 IsExternalAccessPolicyChanged = false,
                 IsHostedVoicemailPolicyChanged = false,
                 IsMobilityPolicyChanged = false,
                 IsTenantDialPlanChanged = false,
                 IsAllowInternationalChanged = false;
            string ChangedAttr = "";


            //check which attributes have been changed
            if (SipAddressTextBox.Text != (String.IsNullOrEmpty(SipAddress) ? SipAddress : SipAddress.Remove(0, 4)))
            {
                ChangedAttr += "\nSip address from " + SipAddress + " to sip:" + SipAddressTextBox.Text + "\n";
                IsSipAddressChanged = true;
            }

            //if (RegistrarPoolTextBlock.Text == RegistrarPool);

            if (LineURITextBox.Text != (String.IsNullOrEmpty(LineURI) ? LineURI : LineURI.Remove(0, 4)))                 
            {
                ChangedAttr += "\nLine URI from " + LineURI + " to " + LineURITextBox.Text + "\n";
                IsLineURIChanged = true;
            }

            if (EnterpriseVoiceEnabledCheckBox.IsChecked != EnterpriseVoiceEnabled)
            {
                ChangedAttr += "\nEV Enabled from " + EnterpriseVoiceEnabled.ToString() + 
                    " to " + EnterpriseVoiceEnabledCheckBox.IsChecked.ToString() + "\n";
                IsEnterpriseVoiceEnabledChanged = true;
            }

            if (HostedVoicemailCheckBox.IsChecked != HostedVoicemail)
            {
                ChangedAttr += "\nHosted VM enabled from " + HostedVoicemail.ToString() +
                    " to " + HostedVoicemailCheckBox.IsChecked.ToString() + "\n";
                IsHostedVoicemailChanged = true;
            }

            if (AllowInternationalCheckBox.IsChecked != AllowInternational)
            {
                ChangedAttr += "\nAllow international from " + AllowInternational.ToString() +
                    " to " + AllowInternationalCheckBox.IsChecked.ToString() + "\n";
                IsAllowInternationalChanged = true;
            }

            //check HybridPstnSite
            if (HybridPstnSiteComboBox.SelectedValue.ToString() != UserPoliciesOfFoundUser["HybridPstnSiteName"].ToString())
            {
                ChangedAttr += "\nPstn Site from " + UserPoliciesOfFoundUser["HybridPstnSiteName"] +
                    " to " + HybridPstnSiteComboBox.SelectedValue + "\n";
                IsHybridPstnSiteNameChanged = true;
            }

            //check TenantDialPlan
            if (TenantDialPlanComboBox.SelectedValue.ToString() != UserPoliciesOfFoundUser["TenantDialPlan"].ToString())
            {
                ChangedAttr += "\nTenant Dial Plan from " + UserPoliciesOfFoundUser["TenantDialPlan"] +
                    " to " + TenantDialPlanComboBox.SelectedValue + "\n";
                IsTenantDialPlanChanged = true;
            }


            //check conf pol
            if (ConferencingPolicyComboBox.SelectedValue.ToString() != UserPoliciesOfFoundUser["ConferencingPolicy"].ToString())
            {
                ChangedAttr += "\nConf policy from " + UserPoliciesOfFoundUser["ConferencingPolicy"] + 
                    " to " + ConferencingPolicyComboBox.SelectedValue + "\n";
                IsConferencingPolicyChanged = true;
            }

            //voice pol
            if (VoicePolicyComboBox.SelectedValue.ToString() != UserPoliciesOfFoundUser["VoicePolicy"].ToString())
            {
                ChangedAttr += "\nVoice policy from " + UserPoliciesOfFoundUser["VoicePolicy"] + 
                    " to " + VoicePolicyComboBox.SelectedValue + "\n";
                IsVoicePolicyChanged = true;
            }

            //ExternalAccess
            if (ExternalAccessPolicyComboBox.SelectedValue.ToString() != UserPoliciesOfFoundUser["ExternalAccessPolicy"].ToString())
            {
                ChangedAttr += "\nExternal Access policy from " + UserPoliciesOfFoundUser["ExternalAccessPolicy"] + 
                    " to " + ExternalAccessPolicyComboBox.SelectedValue + "\n";
                IsExternalAccessPolicyChanged = true;
            }

            //HostedVM
            if (HostedVoicemailPolicyComboBox.SelectedValue.ToString() != UserPoliciesOfFoundUser["HostedVoicemailPolicy"].ToString())
            {
                ChangedAttr += "\nHosted VM policy from " + UserPoliciesOfFoundUser["HostedVoicemailPolicy"] + 
                    " to " + HostedVoicemailPolicyComboBox.SelectedValue + "\n";
                IsHostedVoicemailPolicyChanged = true;
            }

            //mobility pol
            if (MobilityPolicyComboBox.SelectedValue.ToString() != UserPoliciesOfFoundUser["MobilityPolicy"].ToString())
            {
                ChangedAttr += "\nMobility policy from " + UserPoliciesOfFoundUser["MobilityPolicy"] +
                    " to " + MobilityPolicyComboBox.SelectedValue + "\n";
                IsMobilityPolicyChanged = true;
            }

            //start dialog if any of attributes were changed
            if (IsSipAddressChanged ||
               IsLineURIChanged ||
               IsEnterpriseVoiceEnabledChanged ||
               IsHostedVoicemailChanged ||
               IsHybridPstnSiteNameChanged ||
               IsConferencingPolicyChanged ||
               IsVoicePolicyChanged ||
               IsExternalAccessPolicyChanged ||
               IsHostedVoicemailPolicyChanged ||
               IsMobilityPolicyChanged ||
               IsTenantDialPlanChanged ||
               IsAllowInternationalChanged)
            {
                //if confirmed, run set-csuser and\ or grant policies
                if (MessageBox.Show("Do you really want to change following attributes for the "
                    + SipAddress + " user?\n" + ChangedAttr,
                    "Confirm",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information, MessageBoxResult.No) == MessageBoxResult.Yes)
                {
                    //start of set-csuser block

                    if (IsSipAddressChanged)
                    {
                        Collection<PSObject> result = new Collection<PSObject>();
                        result = PSExecute("Set-CsUser -Identity '" + Identity +
                            "'" + " -SipAddress 'sip:" + SipAddressTextBox.Text + "'");

                        if (result != null)
                        {
                            AppendMainText("\nsip address changed\n");
                        }

                        else
                        {
                            AppendMainText("\nThe error occured, please find the error details above, try again\n");
                        }
                    }

                    //run set-csuser once if all 3 parameters were changed
                    if (IsLineURIChanged && IsEnterpriseVoiceEnabledChanged && IsHostedVoicemailChanged)
                    {
                        Collection<PSObject> result = new Collection<PSObject>();
                        string NewLineURI;
                        //if lineuri string is empty, set argument as $null
                        if (String.IsNullOrEmpty(LineURITextBox.Text))
                        {
                            NewLineURI = "$null";
                        }
                        else
                        {
                            NewLineURI = "'tel:" + LineURITextBox.Text + "'";
                        }
                        result = PSExecute("Set-CsUser -Identity '" + Identity + "'" +
                               "-OnPremLineURI " + NewLineURI + " -EnterpriseVoiceEnabled $" + (!EnterpriseVoiceEnabled).ToString() +
                               " -HostedVoicemail $" + (!HostedVoicemail).ToString());

                        if (result != null)
                        {
                            AppendMainText("\nLineURI changed\n" + "\nEV changed\n" + "\nHosted VM enabled changed\n");
                        }

                        else
                        {
                            AppendMainText("\nThe error occured, please find the error details above, try again\n");
                        }

                    }
                
                    //or change one by one
                    else
                    {
                        if (IsLineURIChanged)
                        {
                            string NewLineURI;
                            //if lineuri string is empty, set argument as $null
                            if (String.IsNullOrEmpty(LineURITextBox.Text))
                            {
                                NewLineURI = "$null";
                            }
                            else
                            {
                                NewLineURI = "'tel:" + LineURITextBox.Text + "'";
                            }

                            Collection<PSObject> result = new Collection<PSObject>();
                            result = PSExecute("Set-CsUser -Identity '" + Identity + "'" + "-OnPremLineURI " + NewLineURI);

                            if (result != null)
                            {
                                AppendMainText("\nLineURI changed\n");
                            }

                            else
                            {
                                AppendMainText("\nThe error occured, please find the error details above, try again\n");
                            }
                        }

                        if (IsEnterpriseVoiceEnabledChanged)
                        {
                            Collection<PSObject> result = new Collection<PSObject>();
                            result = PSExecute("Set-CsUser -Identity '" + Identity +
                                "'" + " -EnterpriseVoiceEnabled $" + (!EnterpriseVoiceEnabled).ToString());

                            if (result != null)
                            {
                                AppendMainText("\nEV changed\n");
                            }

                            else
                            {
                                AppendMainText("\nThe error occured, please find the error details above, try again\n");
                            }
                        }

                        if (IsHostedVoicemailChanged)
                        {
                            Collection<PSObject> result = new Collection<PSObject>();
                            result = PSExecute("Set-CsUser -Identity '" + Identity +
                                "'" + " -HostedVoicemail $" + (!HostedVoicemail).ToString());

                            if (result != null)
                            {
                                AppendMainText("\nHosted VM enabled changed\n");
                            }

                            else
                            {
                                AppendMainText("\nThe error occured, please find the error details above, try again\n");
                            }
                        }
                    }
                    //end of set-csuser block

                    if (IsHybridPstnSiteNameChanged)
                    {
                        Collection<PSObject> result = new Collection<PSObject>();
                        if (HybridPstnSiteComboBox.SelectedValue.ToString() != "None")
                            result = PSExecute("Set-CsUserPstnSettings -Identity '" + Identity + "'"
                                + " -HybridPSTNSite '" + HybridPstnSiteComboBox.SelectedValue + "'");
                        else result = PSExecute("Set-CsUserPstnSettings -Identity '" + Identity + "'" + " -HybridPSTNSite $null");
                        if (result != null)
                        {
                            AppendMainText("\nThe user assigned to new Hybrid Pstn Site\n");
                        }

                        else
                        {
                            AppendMainText("\nThe error occured, please find the error details above, try again\n");
                        }
                    }

                    if (IsTenantDialPlanChanged)
                    {
                        Collection<PSObject> result = new Collection<PSObject>();
                        if (TenantDialPlanComboBox.SelectedValue.ToString() != "Global")
                            result = PSExecute("Grant-CsTenantDialPlan -Identity '" + Identity + "'"
                                + " -PolicyName '" + TenantDialPlanComboBox.SelectedValue + "'");
                        else result = PSExecute("Grant-CsTenantDialPlan -Identity '" + Identity + "'" + " -PolicyName $null");
                        if (result != null)
                        {
                            AppendMainText("\nNew Tenant Dial Plan granted\n");
                        }

                        else
                        {
                            AppendMainText("\nThe error occured, please find the error details above, try again\n");
                        }
                    }

                    if (IsConferencingPolicyChanged)
                    {
                        Collection<PSObject> result = new Collection<PSObject>();
                        //for Global and Site policies use $null for automatical assigment
                        if (ConferencingPolicyComboBox.SelectedValue.ToString() != "Global" && !(ConferencingPolicyComboBox.SelectedValue.ToString().StartsWith("Site:")))
                            result = PSExecute("Grant-CsConferencingPolicy -Identity '" + Identity + "'"
                                + " -PolicyName '" + ConferencingPolicyComboBox.SelectedValue + "'");
                        else result = PSExecute("Grant-CsConferencingPolicy -Identity '" + Identity + "'" + " -PolicyName $null");
                        if (result != null)
                        {
                            AppendMainText("\nNew Conferencing Policy granted\n");
                        }

                        else
                        {
                            AppendMainText("\nThe error occured, please find the error details above, try again\n");
                        }
                    }

                    if (IsAllowInternationalChanged)
                    {
                        Collection<PSObject> result = new Collection<PSObject>();

                        if (AllowInternationalCheckBox.IsChecked == true)
                        {
                            result = PSExecute("Grant-CsDialoutPolicy -Identity '" + Identity + "'"
                                + " -PolicyName '" + "tag:DialoutCPCandPSTNInternational" + "'");
                        }
                        else
                        {
                            result = PSExecute("Grant-CsDialoutPolicy -Identity '" + Identity + "'"
                                + " -PolicyName '" + "tag:DialoutCPCandPSTNDomestic" + "'");
                        }

                        if (result != null)
                        {
                            AppendMainText("\nAllow International Changed\n");
                        }

                        else
                        {
                            AppendMainText("\nThe error occured, please find the error details above, try again\n");
                        }
                    }

                    if (IsVoicePolicyChanged)
                    {
                        Collection<PSObject> result = new Collection<PSObject>();
                        //for Global and Site policies use $null for automatical assigment
                        if (VoicePolicyComboBox.SelectedValue.ToString() != "Global" && !(VoicePolicyComboBox.SelectedValue.ToString().StartsWith("Site:")))
                            result = PSExecute("Grant-CsVoicePolicy -Identity '" + Identity + "'"
                                + " -PolicyName '" + VoicePolicyComboBox.SelectedValue + "'");
                        else result = PSExecute("Grant-CsVoicePolicy -Identity '" + Identity + "'" + " -PolicyName $null");
                        if (result != null)
                        {
                            AppendMainText("\nNew Voice Policy granted\n");
                        }

                        else
                        {
                            AppendMainText("\nThe error occured, please find the error details above, try again\n");
                        }
                    }
                    if (IsExternalAccessPolicyChanged)
                    {
                        Collection<PSObject> result = new Collection<PSObject>();
                        //for Global and Site policies use $null for automatical assigment
                        if (ExternalAccessPolicyComboBox.SelectedValue.ToString() != "Global" && !(ExternalAccessPolicyComboBox.SelectedValue.ToString().StartsWith("Site:")))
                            result = PSExecute("Grant-CsExternalAccessPolicy -Identity '" + Identity + "'"
                                + " -PolicyName '" + ExternalAccessPolicyComboBox.SelectedValue + "'");
                        else result = PSExecute("Grant-CsExternalAccessPolicy -Identity '" + Identity + "'" + " -PolicyName $null");
                        if (result != null)
                        {
                            AppendMainText("\nNew External Access Policy granted\n");
                        }

                        else
                        {
                            AppendMainText("\nThe error occured, please find the error details above, try again\n");
                        }
                    }
                    if (IsHostedVoicemailPolicyChanged)
                    {
                        Collection<PSObject> result = new Collection<PSObject>();
                        //for Global and Site policies use $null for automatical assigment
                        if (HostedVoicemailPolicyComboBox.SelectedValue.ToString() != "Global" && !(HostedVoicemailPolicyComboBox.SelectedValue.ToString().StartsWith("Site:")))
                            result = PSExecute("Grant-CsHostedVoicemailPolicy -Identity '" + Identity + "'"
                                + " -PolicyName '" + HostedVoicemailPolicyComboBox.SelectedValue + "'");
                        else result = PSExecute("Grant-CsHostedVoicemailPolicy -Identity '" + Identity + "'" + " -PolicyName $null");
                        if (result != null)
                        {
                            AppendMainText("\nNew Hosted Voicemail Policy granted\n");
                        }

                        else
                        {
                            AppendMainText("\nThe error occured, please find the error details above, try again\n");
                        }
                    }
                    if (IsMobilityPolicyChanged)
                    {
                        Collection<PSObject> result = new Collection<PSObject>();
                        //for Global and Site policies use $null for automatical assigment
                        if (MobilityPolicyComboBox.SelectedValue.ToString() != "Global" && !(MobilityPolicyComboBox.SelectedValue.ToString().StartsWith("Site:")))
                            result = PSExecute("Grant-CsMobilityPolicy -Identity '" + Identity + "'"
                                + " -PolicyName '" + MobilityPolicyComboBox.SelectedValue + "'");
                        else result = PSExecute("Grant-CsMobilityPolicy -Identity '" + Identity + "'" + " -PolicyName $null");
                        if (result != null)
                        {
                            AppendMainText("\nNew Mobility Policy granted\n");
                        }

                        else
                        {
                            AppendMainText("\nThe error occured, please find the error details above, try again\n");
                        }
                    }

                }

            //disable Update Button until next user will be found
            UpdateButton("UpdateUserButton", false, "Search again");

            }

            else
            {
                MessageBox.Show("None of attributes were changed", "Nothing to change", MessageBoxButton.OK,
                MessageBoxImage.Information, MessageBoxResult.OK);
            }
        }    

        //declaration of a delegate for button dispatcher
        private delegate void BUpdater(string ButtonName, bool EnableDisable, string ChangeContent);

        //call this from parrallel thread to update buttons
        private void UpdateButton(string ButtonName, bool EnableDisable, string ChangeContent)
        {
            BUpdater UIUpdater = new BUpdater(ButtonUpdater);
            Dispatcher.BeginInvoke(DispatcherPriority.Send, UIUpdater, ButtonName, EnableDisable, ChangeContent);
        }

        //button dispatcher
        private void ButtonUpdater(string ButtonName, bool EnableDisable, string ChangeContent)
        {
            var myButton = (Button)this.FindName(ButtonName);
            myButton.IsEnabled = EnableDisable;
            if (ChangeContent != "")
            {
                myButton.Content = ChangeContent;
            }       
        }
        
        //get Tenant Dial Plan info
        private void ViewTenantDialPlanButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();

            try
            {
                result = PSExecute("Get-CsTenantDialPlan -Identity '" +
                            TenantDialPlanComboBox.SelectedValue.ToString() + "'").First();
                AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                              + " policy:\n");
                foreach (PSPropertyInfo PSpr in result.Properties)
                {
                    AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                }
            }
            catch (Exception ex)
            {
                AppendMainText("Following exception happened when getting policy info:\n");
                AppendMainText("Exception: " + ex.Message.ToString());
            }
        }

        //get conference policy info
        private void ViewConferencingPolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();

            try
            {
                result = PSExecute("Get-CsConferencingPolicy -Identity '" +
                            ConferencingPolicyComboBox.SelectedValue.ToString() + "'").First();
                AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                              + " policy:\n");
                foreach (PSPropertyInfo PSpr in result.Properties)
                {
                    AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                }
            }
            catch (Exception ex)
            {
                AppendMainText("Following exception happened when getting policy info:\n");
                AppendMainText("Exception: " + ex.Message.ToString());
            }
        }

        //get VoicePolicy info
        private void ViewVoicePolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();
            try
            {
                result = PSExecute("Get-CsVoicePolicy -Identity '" +
                            VoicePolicyComboBox.SelectedValue.ToString() + "'").First();
                AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                              + " policy:\n");
                foreach (PSPropertyInfo PSpr in result.Properties)
                {
                    AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");

                }
            }

            catch (Exception ex)
            {
                AppendMainText("Following exception happened when getting policy info:\n");
                AppendMainText("Exception: " + ex.Message.ToString());
            }
        }
        //get ExternalAccessPolicy info
        private void ViewExternalAccessPolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();

            try
            {
                result = PSExecute("Get-CsExternalAccessPolicy -Identity '" +
                            ExternalAccessPolicyComboBox.SelectedValue.ToString() + "'").First();
                AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                              + " policy:\n");
                foreach (PSPropertyInfo PSpr in result.Properties)
                {
                    AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");              
                }
            }

            catch (Exception ex)
            {
                AppendMainText("Following exception happened when getting policy info:\n");
                AppendMainText("Exception: " + ex.Message.ToString());
            }

        }

        //get HostedVoicemailPolicy info
        private void ViewHostedVoicemailPolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();
            try
            {
                result = PSExecute("Get-CsHostedVoicemailPolicy -Identity '" +
                            HostedVoicemailPolicyComboBox.SelectedValue.ToString() + "'").First();
                AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                              + " policy:\n");

                foreach (PSPropertyInfo PSpr in result.Properties)
                {
                    AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                }
            }

            catch (Exception ex)
            {
                AppendMainText("Following exception happened when getting policy info:\n");
                AppendMainText("Exception: " + ex.Message.ToString());
            }

        }

        private void GetPhoneFromAD_Click(object sender, RoutedEventArgs e)
        {
            LineURITextBox.Text = ADPhone;
        }

        //get MobilityPolicy info
        private void ViewMobilityPolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();
            try
            {
                result = PSExecute("Get-CsMobilityPolicy -Identity '" +
                            MobilityPolicyComboBox.SelectedValue.ToString() + "'").First();
                AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                              + " policy:\n");
                foreach (PSPropertyInfo PSpr in result.Properties)
                {
                    AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                }
            }
        

            catch (Exception ex)
            {
                AppendMainText("Following exception happened when getting policy info:\n");
                AppendMainText("Exception: " + ex.Message.ToString());
            }
        }

        private void AppendMainText(string UpdText)
        {
            //Dispatcher for MainWindow for updating UI
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
            {
                MainText.AppendText(UpdText);
                MainText.ScrollToEnd();
            }
              );
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {

            if (MessageBox.Show("Close Application?", "Exit", MessageBoxButton.YesNo,
                        MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
            {
                e.Cancel = true;
            }
            else
            {
                // dispose the runspace and enable garbage collection
                if (runspace != null)
                {
                runspace.Dispose();
                runspace = null;
                }
            }
        }

        private void ResetPinButton_Click(object sender, RoutedEventArgs e)
        {
                Collection<PSObject> results = new Collection<PSObject>();
                results = PSExecute("Set-CsOnlineDialInConferencingUser -ResetLeaderPin" +
                                                        " -Identity '" + Identity + "'");
                if (results != null)
                {

                AppendMainText("\nThe PIN was reset\n");
                //show only following attributes of a user in main window text box
                string[] UserAttributesToDisplay = { "SipAddress", "ServiceNumber", "TollFreeServiceNumber", "ConferenceId", "BridgeName", "LeaderPin" };
                foreach (string AttrName in UserAttributesToDisplay)
                    {
                    AppendMainText(AttrName + " : " + results.First().Properties[AttrName].Value + "\n");
                    }
                }

                else
                {  
                    AppendMainText("\nThe error occured, please see details above, try again\n");
                }
            
        }

        //restart PS session to prevent a session timeout
        public void OnElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            AppendMainText("\nThe remote PowerShell session expired");

            //disconnect current PS session and disable buttons
            PSDisconnect();

            //AppendMainText("\nNumberOfPSReconnectionsCurrent " + NumberOfPSReconnectionsCurrent);

            if (NumberOfPSReconnectionsCurrent < (NumberOfPSReconnectionsThreshold-1))
            {
                NumberOfPSReconnectionsCurrent++;
                AppendMainText("Re-connecting, wait...\n");
                UpdateButton("ConnectButton", false, "Re-connecting...");
                //create background thread and start PSconnection
                Thread t1 = new Thread(new ThreadStart(PSconnection));
                t1.Start();
            }
            else
            {
                AppendMainText("\nThe number of PS reconnections was too large, close the program and start it again");
            }
        }

        //update UI (fill neccessary attributes in action block (left section)) and fill's combo boxes items with tenant policies 
        private void UpdateUIwithUserAttributes()
        {
            TenantDialPlanComboBox.ItemsSource = TenantDialPlans;
            HybridPstnSiteComboBox.ItemsSource = TenantHybridPSTNSites;
            ConferencingPolicyComboBox.ItemsSource = TenantConferencingPolicies;
            VoicePolicyComboBox.ItemsSource = TenantVoicePolicies;
            ExternalAccessPolicyComboBox.ItemsSource = TenantExternalAccessPolicies;
            HostedVoicemailPolicyComboBox.ItemsSource = TenantHostedVoicemailPolicies;
            MobilityPolicyComboBox.ItemsSource = TenantMobilityPolicies;

            //remove "sip:" prefix if the string isnt empty 
            SipAddressTextBox.Text = String.IsNullOrEmpty(SipAddress) ?  SipAddress : SipAddress.Remove(0, 4);
            RegistrarPoolTextBlock.Text = RegistrarPool;
            //remove "tel:" prefix if the string isnt empty 
            LineURITextBox.Text = String.IsNullOrEmpty(LineURI) ? LineURI : LineURI.Remove(0, 4);
            EnterpriseVoiceEnabledCheckBox.IsChecked = EnterpriseVoiceEnabled;
            HostedVoicemailCheckBox.IsChecked = HostedVoicemail;
            AllowInternationalCheckBox.IsChecked = AllowInternational;

            if (UserPoliciesOfFoundUser != null)
            {
                //select effective policies for current user in the comboboxes
                TenantDialPlanComboBox.SelectedValue = UserPoliciesOfFoundUser["TenantDialPlan"];
                HybridPstnSiteComboBox.SelectedValue = UserPoliciesOfFoundUser["HybridPstnSiteName"];
                ConferencingPolicyComboBox.SelectedValue = UserPoliciesOfFoundUser["ConferencingPolicy"];
                VoicePolicyComboBox.SelectedValue = UserPoliciesOfFoundUser["VoicePolicy"];
                ExternalAccessPolicyComboBox.SelectedValue = UserPoliciesOfFoundUser["ExternalAccessPolicy"];
                HostedVoicemailPolicyComboBox.SelectedValue = UserPoliciesOfFoundUser["HostedVoicemailPolicy"];
                MobilityPolicyComboBox.SelectedValue = UserPoliciesOfFoundUser["MobilityPolicy"];
            }

            //select first value in comboboxes if the user policies are empty
            else
            {
                TenantDialPlanComboBox.SelectedIndex = 0;
                HybridPstnSiteComboBox.SelectedIndex = 0;
                ConferencingPolicyComboBox.SelectedIndex = 0;
                VoicePolicyComboBox.SelectedIndex = 0;
                ExternalAccessPolicyComboBox.SelectedIndex = 0;
                HostedVoicemailPolicyComboBox.SelectedIndex = 0;
                MobilityPolicyComboBox.SelectedIndex = 0;
            }
        }

    }

}


