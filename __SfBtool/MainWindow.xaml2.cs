using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Management.Automation.Runspaces;
using System.Management.Automation.Remoting;
using System.Security;
using System.Threading;
using System.Windows.Threading;

namespace SfBtool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        static Runspace runspace = RunspaceFactory.CreateRunspace();
        static bool ConnectionFlag = false;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        if (ConnectionFlag)
            {
                MainText.AppendText("PS connection has been already established\n");
                return;
            }

            //MainText.AppendText("Connecting to remote PowerShell, wait...\n");
            //ConnectButton.Content = "Working...";
            //this.NavigationService.Refresh();
            Thread t = new Thread(new ThreadStart(Work));
            t.Start();

            //Thread.Sleep(10000);

            string password = "123qwerty=";
            string userName = "lynclab.com\\Andrew";
            System.Uri uri = new Uri("https://192.168.1.134/ocspowershell");
            System.Security.SecureString securePassword = String2SecureString(password);

            PSCredential creds = new PSCredential(userName, securePassword);

            PowerShell powershell = PowerShell.Create();
            PSCommand command = new PSCommand();
            command.AddCommand("New-PSSession");
            command.AddParameter("ConnectionUri", uri);
            command.AddParameter("Credential", creds);
            command.AddParameter("Authentication", "Default");
            //command.AddParameter("RunAsAdministrator");
            PSSessionOption sessionOption = new PSSessionOption();
            sessionOption.SkipCACheck = true;
            sessionOption.SkipCNCheck = true;
            sessionOption.SkipRevocationCheck = true;
            TimeSpan ts = TimeSpan.FromMinutes(60);
            sessionOption.IdleTimeout = ts;
            command.AddParameter("SessionOption", sessionOption);

            powershell.Commands = command;

            try
            {
                // open the remote runspace
                runspace.Open();

                // associate the runspace with powershell
                powershell.Runspace = runspace;

                // invoke the powershell to obtain the results
                Collection<PSSession> result = powershell.Invoke<PSSession>();

                foreach (ErrorRecord current in powershell.Streams.Error)
                {
                    MainText.AppendText("Creating new PSSession error_start\n");
                    MainText.AppendText("Exception: " + current.Exception.ToString() + "\n");
                    MainText.AppendText("Inner Exception: " + current.Exception.InnerException + "\n");
                    MainText.AppendText("Creating new PSSession error_end\n");
                }

                if (result.Count != 1)
                {
                    MainText.AppendText("Couldnt connect to Front End or unexpected number " +
                        "of Remote Runspace connections returned" + "\n");
                    return;
                    //throw new Exception("Unexpected number of Remote Runspace connections returned.");
                }
                
                else
                {
                    MainText.AppendText("New-PSSession created" + "\n");
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
                    MainText.AppendText("Set-Variable error_start\n");
                    MainText.AppendText("Exception: " + current.Exception.ToString() + "\n");
                    MainText.AppendText("Inner Exception: " + current.Exception.InnerException + "\n");
                    MainText.AppendText("Set-Variable error_end\n");
                }

                // 
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
                    MainText.AppendText("Set-ExecutionPolicy error_start\n");
                    MainText.AppendText("Exception: " + current.Exception.ToString() + "\n");
                    MainText.AppendText("Inner Exception: " + current.Exception.InnerException + "\n");
                    MainText.AppendText("Set-ExecutionPolicy error_end\n");
                }

                // First import the cmdlets in the current runspace (using Import-PSSession)
                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("Import-PSSession -Session $ra");
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();

                foreach (ErrorRecord current in powershell.Streams.Error)
                {
                    MainText.AppendText("Import-PSSession error_start\n");
                    MainText.AppendText("Exception: " + current.Exception.ToString() + "\n");
                    MainText.AppendText("Inner Exception: " + current.Exception.InnerException + "\n");
                    MainText.AppendText("Import-PSSession error_end\n");
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
                    MainText.AppendText("Get-CsWindowsService error_start\n");
                    MainText.AppendText("Exception: " + current.Exception.ToString() + "\n");
                    MainText.AppendText("Inner Exception: " + current.Exception.InnerException + "\n");
                    MainText.AppendText("Get-CsWindowsService error_end\n");
                }
                /*
                string servername, serverdomain;

                foreach (PSObject PSresult in results)
                {
                    servernamePSresult.Properties["Name"].Value.ToString() + "\n");
                    MainText.AppendText(PSresult.Properties["sipaddress"].Value.ToString() + "\n");
                }
                */
                if (powershell.Streams.Error.Count == 0)
                {
                    foreach (PSObject PSresult in results)
                    {
                        MainText.AppendText("Successfully connected to " + PSresult.Properties["MachineName"].Value.ToString() +".\n");
                        ConnectionFlag = true;
                        ConnectButton.Content = "Connected";
                        ConnectButton.IsEnabled = false;
                    }
                }
            }

            catch (Exception ex)
            {
                MainText.AppendText("Following exception happened while connecting to remote PS:\n");
                MainText.AppendText("Exception: " + ex.Message.ToString());
            }


            finally
            {
                // dispose the runspace and enable garbage collection
               // runspace.Dispose();
               // runspace = null;

                // Finally dispose the powershell and set all variables to null to free
                // up any resources.
                powershell.Dispose();
                powershell = null;          
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
                MainText.AppendText("Execution called\n");
            }
    }

private static SecureString String2SecureString(string password)
        {
            SecureString remotePassword = new SecureString();
            for (int i = 0; i < password.Length; i++)
                remotePassword.AppendChar(password[i]);

            return remotePassword;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Collection<PSObject> results = new Collection<PSObject>();
            results = PSExecute(InputCommand.Text);
            if (results != null)
            {
                foreach (PSObject PSresult in results)
                {
                    //MainText.AppendText(PSresult.Methods.ToList().ToString());
                    //MainText.AppendText(PSresult.Properties["Name"].Value.ToString() + "\n");
                    MainText.AppendText(PSresult.ToString() + "\n");
                }
            }
            else
            {
                MainText.AppendText("Null PS Object was returned\n");
            }
        }

        private void TestButton_Click(object sender, RoutedEventArgs e)
        {
            MainText.AppendText(runspace.RunspaceStateInfo.ToString());
           // Availability
        }


//>>>>>>>>>>>>>>>>>>>>
        private void Work()
        {
            //   for (int i = 0; i <= 50; i++)
            // {
                Updater uiUpdater = new Updater(UpdateUI);
                Dispatcher.BeginInvoke(DispatcherPriority.Send, uiUpdater,"Wait...");
            // Thread.Sleep(500);
            //   }
        }

        private delegate void Updater(string UI);

        private void UpdateUI(string ButtonText)
        {
            ConnectButton.Content = ButtonText;
        }
        //>>>>>>>>>>>>>>>>>>>>
    }
}


