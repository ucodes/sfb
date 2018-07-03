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
        string SipAddress, RegistrarPool, LineURI;

        Hashtable UserPolicies = new Hashtable();
        bool EnterpriseVoiceEnabled;

        string password;
        string userName;

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


            /*
            if (ConnectionFlag)
            {
                MainText.AppendText("PS connection has been already established\n");
                return;
            }
            */

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

            System.Uri uri = new Uri("https://192.168.1.134/ocspowershell");
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
            TimeSpan ts = TimeSpan.FromMinutes(1);
            sessionOption.IdleTimeout = ts;
            sessionOption.CancelTimeout = ts;
            sessionOption.OperationTimeout = ts;
            sessionOption.OpenTimeout = ts;
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
                }

                //one PS sessions is expected
                if (result.Count != 1)
                {
                    returnstring += "Couldnt connect to Front End or unexpected number " +
                        "of Remote Runspace connections returned" + "\n";
                   // return returnstring;
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

            List<string>  _HostedVoicemailPolicy = new List<string>();
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
                    MainText.AppendText("PSExecute error start\n");
                    MainText.AppendText("Exception: " + current.Exception.ToString() + "\n");
                    MainText.AppendText("Inner Exception: " + current.Exception.InnerException + "\n");
                    MainText.AppendText("PSExecute error end\n");
                }

                return results;

            /*
            foreach (PSObject PSresult in results)
            {
                MainText.AppendText(PSresult.Methods.ToList().ToString());
                MainText.AppendText(PSresult.Properties["Name"].Value.ToString() + "\n");
                MainText.AppendText(PSresult.Properties["sipaddress"].Value.ToString() + "\n");
            }
            */
        }

            catch (Exception ex)
            {
                MainText.AppendText("Following exception happened when calling PS command:\n");
                MainText.AppendText("Exception: " + ex.Message.ToString());
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
               // ConnectButton.Content = "Connected";
               // MainText.AppendText("Execution called\n");
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
            SipAddress = "";
            RegistrarPool = "";
            LineURI = "";
            EnterpriseVoiceEnabled = false;

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
            //check whether more than one user was found
            if (results.Count > 1)
            {
                AppendMainText("\nFollowing users were found, please type exact SIP address of a user" + "\n");
                //MainText.AppendText("Following users were found, please type exact SIP address of a user" + "\n");

                foreach (PSObject PSresult in results)
                {
                    AppendMainText(PSresult.Properties["SipAddress"].Value.ToString() + "\n");
                    //MainText.AppendText(PSresult.Properties["SipAddress"].Value.ToString() + "\n");
                }
            }
            //one user is expected
            else if (results.Count == 1)
            {

                //show all user info in the main box
                AppendMainText("\nAll attributes of found user:" + "\n");
                //MainText.AppendText("All attributes of found user:" + "\n");
                foreach (PSPropertyInfo PSpr in results.First().Properties)
                {
                    AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                    //MainText.AppendText(PSpr.Name + " " + PSpr.Value + "\n");
                }

                //fill neccessary attributes in action block (left section)

                SipAddress = results.First().Properties["SipAddress"].Value.ToString();
                RegistrarPool = results.First().Properties["RegistrarPool"].Value.ToString();
                if (results.First().Properties["EnterpriseVoiceEnabled"].Value.ToString() == "True")
                {
                    EnterpriseVoiceEnabled = true;
                }
                LineURI = results.First().Properties["LineURI"].Value.ToString();

                //get effective user policies
                results = PSExecute("Get-CsEffectivePolicy " + SipAddress);
                UserPolicies.Add("ConferencingPolicy", results.First().Properties["ConferencingPolicy"].Value.ToString());
                UserPolicies.Add("VoicePolicy", results.First().Properties["VoicePolicy"].Value.ToString());
                UserPolicies.Add("ExternalAccessPolicy", results.First().Properties["ExternalAccessPolicy"].Value.ToString());
                UserPolicies.Add("HostedVoicemailPolicy", results.First().Properties["HostedVoicemailPolicy"].Value.ToString());
                UserPolicies.Add("MobilityPolicy", results.First().Properties["MobilityPolicy"].Value.ToString());

                //change site policy name from "Site:siteID"  to "Site:siteName"
                List<string> _keys = new List<string>(UserPolicies.Keys.Cast<string>());
                //_keys = UserPolicies;

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

                //update UI
                UpdateUIwithUserAttributes();

                //enable functional buttons
                UpdateButton("ViewConferencingPolicyButton", true, "");
                UpdateButton("ViewExternalAccessPolicyButton", true, "");
                UpdateButton("ViewHostedVoicemailPolicyButton", true, "");
                UpdateButton("ViewMobilityPolicyButton", true, "");
                UpdateButton("ViewVoicePolicyButton", true, "");
                UpdateButton("UpdateUserButton", true, "");
                UpdateButton("ResetPinButton", true, "");

            }

            else
            {
                AppendMainText("\nNo users were found\n");
            }
        }

        private void UpdateUserButton_Click(object sender, RoutedEventArgs e)
        {
            
            MessageBox.Show("Do you really want to change following attributes for the " + SipAddress + " user?", 
                "Confirm", MessageBoxButton.OKCancel,
                MessageBoxImage.Information, MessageBoxResult.Cancel);
        }

        /*
                //>>>>>>>>>>>>>>>>>>>>
                private void UpdateButton(string updatestring)
                {

                    //  Updater uiUpdater = new Updater(UpdateUI);
                    // Dispatcher.BeginInvoke(DispatcherPriority.Send, uiUpdater, "Wait...");

                    //Dispatcher for MainWindow for updating UI
                    Dispatcher.BeginInvoke(DispatcherPriority.Normal,(ThreadStart)delegate ()
            {
                MainText.AppendText(updatestring);
                //UpdateUI();
                ConnectButton.Content = "Connected";
                var myTextBox = (TextBox)this.FindName("InputCommand");
                myTextBox.AppendText("ok!");
                //PSconnection();
            }
              );


                }
                */

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
                ConnectButton.Content = ChangeContent;
            }       
        }

        //get conference policy info
        private void ViewConferencingPolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();
            result = PSExecute("Get-CsConferencingPolicy -Identity " +
                        ConferencingPolicyComboBox.SelectedValue.ToString()).First();
            AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                          + " policy:\n");
            //MainText.AppendText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
            //            + "policy:\n");
            foreach (PSPropertyInfo PSpr in result.Properties)
            {
                AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                //MainText.AppendText(PSpr.Name + " " + PSpr.Value + "\n");
            }
        }

        //get VoicePolicy info
        private void ViewVoicePolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();
            result = PSExecute("Get-CsVoicePolicy -Identity " +
                        VoicePolicyComboBox.SelectedValue.ToString()).First();
            AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                          + " policy:\n");
            //MainText.AppendText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
            //            + "policy:\n");
            foreach (PSPropertyInfo PSpr in result.Properties)
            {
                AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                //MainText.AppendText(PSpr.Name + " " + PSpr.Value + "\n");
            }
        }
        //get ExternalAccessPolicy info
        private void ViewExternalAccessPolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();
            result = PSExecute("Get-CsExternalAccessPolicy -Identity " +
                        ExternalAccessPolicyComboBox.SelectedValue.ToString()).First();
            AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                          + " policy:\n");
            //MainText.AppendText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
            //            + "policy:\n");
            foreach (PSPropertyInfo PSpr in result.Properties)
            {
                AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                //MainText.AppendText(PSpr.Name + " " + PSpr.Value + "\n");
            }
        }
        //get HostedVoicemailPolicy info
        private void ViewHostedVoicemailPolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();
            result = PSExecute("Get-CsHostedVoicemailPolicy -Identity " +
                        HostedVoicemailPolicyComboBox.SelectedValue.ToString()).First();
            AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                          + " policy:\n");
            //MainText.AppendText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
            //            + "policy:\n");
            foreach (PSPropertyInfo PSpr in result.Properties)
            {
                AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                //MainText.AppendText(PSpr.Name + " " + PSpr.Value + "\n");
            }
        }
        //get MobilityPolicy info
        private void ViewMobilityPolicyButton_Click(object sender, RoutedEventArgs e)
        {
            PSObject result = new PSObject();
            result = PSExecute("Get-CsMobilityPolicy -Identity " +
                        MobilityPolicyComboBox.SelectedValue.ToString()).First();
            AppendMainText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
                          + " policy:\n");
            //MainText.AppendText("\nAll properties of " + result.Properties["Identity"].Value.ToString()
            //            + "policy:\n");
            foreach (PSPropertyInfo PSpr in result.Properties)
            {
                AppendMainText(PSpr.Name + " " + PSpr.Value + "\n");
                //MainText.AppendText(PSpr.Name + " " + PSpr.Value + "\n");
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

        //run command on a remote server to prevent a session timeout
        private void OnElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            PSExecute("get-cssipdomain");
           // AppendMainText("timer called");
            Idletimer.Start(); // Restart timer
        }


        private void UpdateUIwithUserAttributes()
        {

            SipAddressTextBox.Text = SipAddress;
            RegistrarPoolTextBlock.Text = RegistrarPool;
            LineURITextBox.Text = LineURI;
            EnterpriseVoiceEnabledCheckBox.IsChecked = EnterpriseVoiceEnabled;

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


