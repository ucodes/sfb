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
        System.Timers.Timer Idletimer = new System.Timers.Timer(45 * 1000);

        static bool ConnectionFlag = false;

        //user attributes
        string Identity, SipAddress, RegistrarPool, LineURI;

        Hashtable UserPolicies = new Hashtable();
        bool EnterpriseVoiceEnabled, HostedVoicemail;

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

            string returnstring="";

            System.Uri uri = new Uri("https://localhost:12346/ocspowershell");
            System.Security.SecureString securePassword = String2SecureString(password);

            PSCredential creds = new PSCredential(userName, securePassword);

            runspace = RunspaceFactory.CreateRunspace();

            PowerShell powershell = PowerShell.Create();
            PSCommand command = new PSCommand();
            command.AddCommand("New-PSSession");
            command.AddParameter("ConnectionUri", uri);
            command.AddParameter("Credential", creds);
            command.AddParameter("Authentication", "Default");
            PSSessionOption sessionOption = new PSSessionOption();
            sessionOption.SkipCACheck = true;
            sessionOption.SkipCNCheck = true;
            sessionOption.SkipRevocationCheck = true;
            TimeSpan ts = TimeSpan.FromMinutes(5);
            sessionOption.IdleTimeout = ts;
            sessionOption.CancelTimeout = ts;
            sessionOption.OperationTimeout = ts;
            sessionOption.OpenTimeout = TimeSpan.FromMinutes(1);
            command.AddParameter("SessionOption", sessionOption);

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
                    returnstring += "Couldnt connect to Front End or unexpected number " +
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
                command.AddScript("Get-CsWindowsService RTCSRV | Select-Object *");
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
                            results.First().Properties["MachineName"].Value.ToString() + " .Search for a user and " +
                                    "change required attributes" + ".\n");
                    ConnectionFlag = true;
                    //update buttons
                    UpdateButton("ConnectButton", false, "Connected");
                    UpdateButton("SearchButton", true, "");

                    //start timer to prevent a session timeout
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
                    UpdateButton("ConnectButton", true, "Connect");
                    UpdateButton("SearchButton", false, "");

                    // dispose the runspace and enable garbage collection if connection was unsuccessfull
                    runspace.Dispose();
                    runspace = null;
                }

            }
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
                List<string> _ConferencingPolicy = new List<string>();
                results = PSExecute("Get-CsConferencingPolicy");
                foreach (PSObject PSresult in results)
                {
                    _ConferencingPolicy.Add(PSresult.Properties["Identity"].Value.ToString());
                }

                ConferencingPolicyComboBox.ItemsSource = _ConferencingPolicy;

                List<string> _VoicePolicy = new List<string>();
                results = PSExecute("Get-CsVoicePolicy");
                foreach (PSObject PSresult in results)
                {
                    _VoicePolicy.Add(PSresult.Properties["Identity"].Value.ToString());
                }

                VoicePolicyComboBox.ItemsSource = _VoicePolicy;

                List<string> _ExternalAccessPolicy = new List<string>();
                results = PSExecute("Get-CsExternalAccessPolicy");
                foreach (PSObject PSresult in results)
                {
                    _ExternalAccessPolicy.Add(PSresult.Properties["Identity"].Value.ToString());
                }

                ExternalAccessPolicyComboBox.ItemsSource = _ExternalAccessPolicy;

                List<string> _HostedVoicemailPolicy = new List<string>();
                results = PSExecute("Get-CsHostedVoicemailPolicy");
                foreach (PSObject PSresult in results)
                {
                    _HostedVoicemailPolicy.Add(PSresult.Properties["Identity"].Value.ToString());
                }

                HostedVoicemailPolicyComboBox.ItemsSource = _HostedVoicemailPolicy;

                List<string> _MobilityPolicy = new List<string>();
                results = PSExecute("Get-CsMobilityPolicy");
                foreach (PSObject PSresult in results)
                {
                    _MobilityPolicy.Add(PSresult.Properties["Identity"].Value.ToString());
                }

                MobilityPolicyComboBox.ItemsSource = _MobilityPolicy;
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
                //powershell.Dispose();
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
            //clean attributes from previous user
            UserPolicies.Clear();
            Identity = "";
            SipAddress = "";
            RegistrarPool = "";
            LineURI = "";
            EnterpriseVoiceEnabled = false;
            HostedVoicemail = false;

            //update UI
            UpdateUIwithUserAttributes();

            //disable functional buttons
            UpdateButton("ViewConferencingPolicyButton", false, "");
            UpdateButton("ViewExternalAccessPolicyButton", false, "");
            UpdateButton("ViewHostedVoicemailPolicyButton", false, "");
            UpdateButton("ViewMobilityPolicyButton", false, "");
            UpdateButton("ViewVoicePolicyButton", false, "");
            UpdateButton("UpdateUserButton", false, "");
            UpdateButton("ResetPinButton", false, "");
            
            //retrive policies info to fill combo boxes items  
            GetPolicyInfo();

            Collection<PSObject> results = new Collection<PSObject>();
            results = PSExecute("Get-CsUser -Filter{SipAddress -like \"*" + 
                                    SearchUserTextBox.Text + "*\"}");

            try
            {

                //check whether more than one user was found
                if (results.Count > 1)
                {
                    AppendMainText("\nFollowing users were found, please type exact SIP address of a user" + "\n");

                    foreach (PSObject PSresult in results)
                    {
                        AppendMainText(PSresult.Properties["SipAddress"].Value.ToString() + "\n");
                    }
                }
                //one user is expected
                else if (results.Count == 1)
                {

                    //show all user info in the main box
                    AppendMainText("\nAll attributes of found user:" + "\n");
                    foreach (PSPropertyInfo PSpr in results.First().Properties)
                    {
                        AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                    }

                    //get user info
                    Identity = results.First().Properties["Identity"].Value.ToString();
                    SipAddress = results.First().Properties["SipAddress"].Value.ToString();
                    RegistrarPool = results.First().Properties["RegistrarPool"].Value.ToString();
                    if (results.First().Properties["EnterpriseVoiceEnabled"].Value.ToString() == "True")
                    {
                        EnterpriseVoiceEnabled = true;
                    }

                    if (results.First().Properties["HostedVoiceMail"].Value != null)
                    {
                        if (results.First().Properties["HostedVoiceMail"].Value.ToString() == "True")
                        {
                            HostedVoicemail = true;
                        }
                    }
                

                    //check - if lineuri is null set up an empty string 
                    if (results.First().Properties["LineURI"].Value == null)
                    {
                        LineURI = "";
                    }
                    else
                    {
                        LineURI = results.First().Properties["LineURI"].Value.ToString();
                    }              

                    //get effective user policies
                    results = PSExecute("Get-CsEffectivePolicy -Identity '" + Identity + "'");
                    UserPolicies.Add("ConferencingPolicy", results.First().Properties["ConferencingPolicy"].Value.ToString());
                    UserPolicies.Add("VoicePolicy", results.First().Properties["VoicePolicy"].Value.ToString());
                    UserPolicies.Add("ExternalAccessPolicy", results.First().Properties["ExternalAccessPolicy"].Value.ToString());
                    UserPolicies.Add("HostedVoicemailPolicy", results.First().Properties["HostedVoicemailPolicy"].Value.ToString());
                    UserPolicies.Add("MobilityPolicy", results.First().Properties["MobilityPolicy"].Value.ToString());

                    //change site policy name from "Site:siteID"  to "Site:siteName"
                    List<string> _keys = new List<string>(UserPolicies.Keys.Cast<string>());

                    foreach (string policy in _keys)
                    {
                        if (UserPolicies[policy].ToString().StartsWith("Site:"))
                        {
                            //get site name by site id taken from effective policy
                            results = PSExecute("Get-CsSite | Where-Object SiteId -eq \"" +
                                                    UserPolicies[policy].ToString().Replace("Site:", "") + "\"");
                            UserPolicies[policy] = results.First().Properties["Identity"].Value.ToString();
                        }
                    }

                    //update UI (fill neccessary attributes in action block (left section))
                    UpdateUIwithUserAttributes();

                    //enable functional buttons
                    UpdateButton("ViewConferencingPolicyButton", true, "");
                    UpdateButton("ViewExternalAccessPolicyButton", true, "");
                    UpdateButton("ViewHostedVoicemailPolicyButton", true, "");
                    UpdateButton("ViewMobilityPolicyButton", true, "");
                    UpdateButton("ViewVoicePolicyButton", true, "");
                    UpdateButton("UpdateUserButton", true, "Update");
                    UpdateButton("ResetPinButton", true, "");

                }

                else
                {
                    AppendMainText("\nNo users were found\n");
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
                 //   IsRegistrarPoolChanged = false, 
                 IsLineURIChanged = false,
                 IsEnterpriseVoiceEnabledChanged = false,
                 IsHostedVoicemailChanged = false,
                 IsConferencingPolicyChanged = false,
                 IsVoicePolicyChanged = false,
                 IsExternalAccessPolicyChanged = false,
                 IsHostedVoicemailPolicyChanged = false,
                 IsMobilityPolicyChanged = false;
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
                ChangedAttr += "\nLine URI from " + LineURI + " to tel:" + LineURITextBox.Text + "\n";
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

            //check conf pol
            if (ConferencingPolicyComboBox.SelectedValue != UserPolicies["ConferencingPolicy"])
            {
                ChangedAttr += "\nConf policy from " + UserPolicies["ConferencingPolicy"] + 
                    " to " + ConferencingPolicyComboBox.SelectedValue + "\n";
                IsConferencingPolicyChanged = true;
            }

            //voice pol
            if (VoicePolicyComboBox.SelectedValue != UserPolicies["VoicePolicy"])
            {
                ChangedAttr += "\nVoice policy from " + UserPolicies["VoicePolicy"] + 
                    " to " + VoicePolicyComboBox.SelectedValue + "\n";
                IsVoicePolicyChanged = true;
            }

            //ExternalAccess
            if (ExternalAccessPolicyComboBox.SelectedValue != UserPolicies["ExternalAccessPolicy"])
            {
                ChangedAttr += "\nExternal Access policy from " + UserPolicies["ExternalAccessPolicy"] + 
                    " to " + ExternalAccessPolicyComboBox.SelectedValue + "\n";
                IsExternalAccessPolicyChanged = true;
            }

            //HostedVM
            if (HostedVoicemailPolicyComboBox.SelectedValue != UserPolicies["HostedVoicemailPolicy"])
            {
                ChangedAttr += "\nHosted VM policy from " + UserPolicies["HostedVoicemailPolicy"] + 
                    " to " + HostedVoicemailPolicyComboBox.SelectedValue + "\n";
                IsHostedVoicemailPolicyChanged = true;
            }

            //mobility pol
            if (MobilityPolicyComboBox.SelectedValue != UserPolicies["MobilityPolicy"])
            {
                ChangedAttr += "\nMobility policy from " + UserPolicies["MobilityPolicy"] +
                    " to " + MobilityPolicyComboBox.SelectedValue + "\n";
                IsMobilityPolicyChanged = true;
            }

            //start dialog if any of attributes were changed
            if (IsSipAddressChanged ||
               IsLineURIChanged ||
               IsEnterpriseVoiceEnabledChanged ||
               IsHostedVoicemailChanged ||
               IsConferencingPolicyChanged ||
               IsVoicePolicyChanged ||
               IsExternalAccessPolicyChanged ||
               IsHostedVoicemailPolicyChanged ||
               IsMobilityPolicyChanged)
            {
                //if confirmed, run set-csuser and\ or grant policies
                if (MessageBox.Show("Do you really want to change following attributes for the "
                    + SipAddress + " user? (if Global or Site: policies are chosen, the user " +
                    "will be granted with automatic policy)\n" + ChangedAttr,
                    "Confirm",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information, MessageBoxResult.No) == MessageBoxResult.Yes)
                {
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
                        result = PSExecute("Set-CsUser -Identity '" + Identity + "'" + "-LineURI " + NewLineURI);

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

        //get conference policy info
        private void ViewConferencingPolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();

            try
            {
                result = PSExecute("Get-CsConferencingPolicy -Identity " +
                            ConferencingPolicyComboBox.SelectedValue.ToString()).First();
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
                result = PSExecute("Get-CsVoicePolicy -Identity " +
                            VoicePolicyComboBox.SelectedValue.ToString()).First();
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
                result = PSExecute("Get-CsExternalAccessPolicy -Identity " +
                            ExternalAccessPolicyComboBox.SelectedValue.ToString()).First();
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
                result = PSExecute("Get-CsHostedVoicemailPolicy -Identity " +
                            HostedVoicemailPolicyComboBox.SelectedValue.ToString()).First();
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
        //get MobilityPolicy info
        private void ViewMobilityPolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();
            try
            {
                result = PSExecute("Get-CsMobilityPolicy -Identity " +
                            MobilityPolicyComboBox.SelectedValue.ToString()).First();
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

            EnterPIN pinwindow = new EnterPIN();
            pinwindow.Owner = this;
            pinwindow.ShowDialog();

            //check whether Pin was entered and isnt null
            if (pinwindow.PinEntered && !String.IsNullOrEmpty(pinwindow.Pin))
            {
                Collection<PSObject> result = new Collection<PSObject>();
                result = PSExecute("Set-CsClientPin -Pin '" + pinwindow.Pin + "'" + " -Identity '" + Identity + "'");
                if (result != null)
                     {
                         AppendMainText(String.Format("\nThe pin reset for {0}\nPin: {1}\nPinReset: {2}", result.First().Properties["Identity"].Value.ToString(),
                         result.First().Properties["Pin"].Value.ToString(),
                         result.First().Properties["PinReset"].Value.ToString()));
                     }

                else
                     {  
                        AppendMainText("\nThe error occured, please see details above, try again\n");
                     }
            }
        }

        //run command on a remote server to prevent a session timeout
        private void OnElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            PSExecute("get-cssipdomain");
            Idletimer.Start(); // Restart timer
        }


        private void UpdateUIwithUserAttributes()
        {
            //remove "sip:" prefix if the string isnt empty 
            SipAddressTextBox.Text = String.IsNullOrEmpty(SipAddress) ?  SipAddress : SipAddress.Remove(0, 4);
            RegistrarPoolTextBlock.Text = RegistrarPool;
            //remove "tel:" prefix if the string isnt empty 
            LineURITextBox.Text = String.IsNullOrEmpty(LineURI) ? LineURI : LineURI.Remove(0, 4);
            EnterpriseVoiceEnabledCheckBox.IsChecked = EnterpriseVoiceEnabled;
            HostedVoicemailCheckBox.IsChecked = HostedVoicemail;

            if (UserPolicies != null)
            {
                //select effective policies for current user in the comboboxes
                ConferencingPolicyComboBox.SelectedValue = UserPolicies["ConferencingPolicy"];
                VoicePolicyComboBox.SelectedValue = UserPolicies["VoicePolicy"];
                ExternalAccessPolicyComboBox.SelectedValue = UserPolicies["ExternalAccessPolicy"];
                HostedVoicemailPolicyComboBox.SelectedValue = UserPolicies["HostedVoicemailPolicy"];
                MobilityPolicyComboBox.SelectedValue = UserPolicies["MobilityPolicy"];
            }

            //select first value in comboboxes if the user policies are empty
            else
            {
                ConferencingPolicyComboBox.SelectedIndex = 0;
                VoicePolicyComboBox.SelectedIndex = 0;
                ExternalAccessPolicyComboBox.SelectedIndex = 0;
                HostedVoicemailPolicyComboBox.SelectedIndex = 0;
                MobilityPolicyComboBox.SelectedIndex = 0;
            }
        }

    }

}


