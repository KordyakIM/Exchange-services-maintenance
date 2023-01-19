using Microsoft.SharePoint;
using System;
using System.Management.Automation;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using System.Text;
using System.Text.RegularExpressions;
using System.Net.NetworkInformation;
using System.Collections.Generic;
using System.Web;
using System.IO;

namespace ExchangeServices.VisualWebPart1
{
    public partial class VisualWebPart1UserControl : UserControl
    {
        string script;
        PowerShell ps;
        bool pingable;
        bool boolServers_array01;
        bool boolServers_array02;
        string input;
        string stringPathLog;
        Ping pinger;
        protected void Page_Load(object sender, EventArgs e)
        {
            input = "Passcode";
            stringPathLog = @"C:\path\ExchangeServices.log";
            pingable = false;
            pinger = null;
            pinger = new Ping();
            Page.Server.ScriptTimeout = 6400; // specify the timeout to 6400 seconds
        }
        protected void Start_click(object sender, EventArgs e)
        {
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                if (TextBox1.Text == input)
                {
                    ResultBox.Text = string.Empty;
                    var builder = new StringBuilder();
                    List<string> arrayMailboxServers = new List<string>();
                    List<string> arrayServices = new List<string>();
                    for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                    {
                        if (CheckBoxList_array01.Items[i].Selected)
                        {
                            arrayMailboxServers.Add(CheckBoxList_array01.Items[i].Text);
                        }
                    }
                    for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                    {
                        if (CheckBoxList_array02.Items[i].Selected)
                        {
                            arrayMailboxServers.Add(CheckBoxList_array02.Items[i].Text);
                        }
                    }
                    for (int i = 0; i < CheckBoxList_Services.Items.Count; i++)
                    {
                        if (CheckBoxList_Services.Items[i].Selected)
                        {
                            string stringReplaceService = Regex.Replace(CheckBoxList_Services.Items[i].Text, @".*[(]", "");
                            string stringReplaceService1 = Regex.Replace(stringReplaceService, @"[)]", "");
                            arrayServices.Add(stringReplaceService1);
                        }
                    }
                    if (arrayMailboxServers.Count > 0)
                    {
                        if (arrayServices.Count > 0)
                        {
                            ResultBox.Text = "+++++ START +++++" + "\r\n";
                            ResultBox.Text += "DateTime start: " + DateTime.Now.ToString() + "\r\n";
                            ResultBox.Text += "Username: " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + "\r\n";
                            foreach (object objectServer in arrayMailboxServers)
                            {
                                string stringMailboxServer = (string)objectServer;
                                foreach (object objectService in arrayServices)
                                {
                                    string stringService = (string)objectService;
                                    PingReply reply = pinger.Send(stringMailboxServer);
                                    pingable = reply.Status == IPStatus.Success;
                                    if (pingable)
                                    {
                                        script = @"
                                        $password = ConvertTo-SecureString 'Password' -AsPlainText -Force
                                        $cred = New-Object System.Management.Automation.PSCredential('Username', $password)
                                        Invoke-Command -ComputerName '" + stringMailboxServer + @"' -ScriptBlock {Start-Service '" + stringService + @"'} -Credential $cred
                                        if($?){
                                            Write-Output 'Successful STARTED service - " + stringService + @" on server - " + stringMailboxServer + @"!'
                                            Invoke-Command -ComputerName '" + stringMailboxServer + @"' -ScriptBlock {Get-Service -Name '" + stringService + @"'} -Credential $cred | select PSComputerName,name,status
                                        }else{
                                            Write-Output 'ERROR'; $error; break
                                        }
                                        ";
                                        using (ps = PowerShell.Create())
                                        {
                                            PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                            ps.AddScript(script);
                                            IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                            ps.EndInvoke(invocation);
                                            ps.Stop();
                                            foreach (PSObject outp in output)
                                            {
                                                if (!outp.BaseObject.ToString().Contains("tmp"))
                                                {
                                                    ResultBox.Text += outp.ToString() + "\r\n";
                                                }
                                            }
                                        }
                                        ResultBox.Text += builder.ToString();
                                    }
                                    else { ResultBox.Text += "\r\n" + "No ping -> " + stringMailboxServer; }
                                }
                            }
                            for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                            {
                                if (CheckBoxList_array01.Items[i].Selected)
                                {
                                    CheckBoxList_array01.Items[i].Selected = false;
                                }
                            }
                            for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                            {
                                if (CheckBoxList_array02.Items[i].Selected)
                                {
                                    CheckBoxList_array02.Items[i].Selected = false;
                                }
                            }
                            for (int i = 0; i < CheckBoxList_Services.Items.Count; i++)
                            {
                                if (CheckBoxList_Services.Items[i].Selected)
                                {
                                    CheckBoxList_Services.Items[i].Selected = false;
                                }
                            }
                            ResultBox.Text += "DateTime end: " + DateTime.Now.ToString() + "\r\n";
                            ResultBox.Text += "+++++ END +++++";
                            File.AppendAllText(stringPathLog, ResultBox.Text);
                        }
                        else { ResultBox.Text += "Please, select service."; }
                    }
                    else { ResultBox.Text += "Please, select Mailbox server."; }
                }
                else { ResultBox.Text = "Wrong Input Code!"; }
            }
        }
        protected void Stop_click(object sender, EventArgs e)
        {
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                if (TextBox1.Text == input)
                {
                    ResultBox.Text = string.Empty;
                    var builder = new StringBuilder();
                    List<string> arrayMailboxServers = new List<string>();
                    //List<string> arrayMailboxServers_array02 = new List<string>();
                    List<string> arrayServices = new List<string>();
                    for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                    {
                        if (CheckBoxList_array01.Items[i].Selected)
                        {
                            arrayMailboxServers.Add(CheckBoxList_array01.Items[i].Text);
                        }
                    }
                    for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                    {
                        if (CheckBoxList_array02.Items[i].Selected)
                        {
                            arrayMailboxServers.Add(CheckBoxList_array02.Items[i].Text);
                        }
                    }
                    for (int i = 0; i < CheckBoxList_Services.Items.Count; i++)
                    {
                        if (CheckBoxList_Services.Items[i].Selected)
                        {
                            string stringReplaceService = Regex.Replace(CheckBoxList_Services.Items[i].Text, @".*[(]", "");
                            string stringReplaceService1 = Regex.Replace(stringReplaceService, @"[)]", "");
                            arrayServices.Add(stringReplaceService1);
                        }
                    }
                    if (arrayMailboxServers.Count > 0)
                    {
                        string stringMailboxServers = string.Join(" ", arrayMailboxServers);
                        boolServers_array01 = Regex.IsMatch(stringMailboxServers, @"nr", RegexOptions.IgnoreCase);
                        boolServers_array02 = Regex.IsMatch(stringMailboxServers, @"nt", RegexOptions.IgnoreCase);
                        if ((boolServers_array01 & !boolServers_array02) || (!boolServers_array01 & boolServers_array02))
                        {
                            if (arrayServices.Count > 0)
                            {
                                ResultBox.Text = "+++++ START +++++" + "\r\n";
                                ResultBox.Text += "DateTime start: " + DateTime.Now.ToString() + "\r\n";
                                ResultBox.Text += "Username: " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + "\r\n";
                                foreach (object objectServer in arrayMailboxServers)
                                {
                                    string stringMailboxServer = (string)objectServer;
                                    foreach (object objectService in arrayServices)
                                    {
                                        string stringService = (string)objectService;
                                        PingReply reply = pinger.Send(stringMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        if (pingable)
                                        {
                                            script = @"
                                            $password = ConvertTo-SecureString 'Password' -AsPlainText -Force
                                            $cred = New-Object System.Management.Automation.PSCredential('Username', $password)
                                            Invoke-Command -ComputerName '" + stringMailboxServer + @"' -ScriptBlock {Stop-Service '" + stringService + @"'} -Credential $cred
                                            if($?){
                                                Write-Output 'Successful STOPPED service - " + stringService + @" on server - " + stringMailboxServer + @"!'
                                                Invoke-Command -ComputerName '" + stringMailboxServer + @"' -ScriptBlock {Get-Service -Name '" + stringService + @"'} -Credential $cred | select PSComputerName,name,status
                                            }else{
                                                Write-Output 'ERROR'; $error; break
                                            }
                                            ";
                                            using (ps = PowerShell.Create())
                                            {
                                                PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                                ps.AddScript(script);
                                                IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                                ps.EndInvoke(invocation);
                                                ps.Stop();
                                                foreach (PSObject outp in output)
                                                {
                                                    if (!outp.BaseObject.ToString().Contains("tmp"))
                                                    {
                                                        ResultBox.Text += outp.ToString() + "\r\n";
                                                    }
                                                }
                                            }
                                            ResultBox.Text += builder.ToString();
                                        }
                                        else { ResultBox.Text += "\r\n" + "No ping -> " + stringMailboxServer; }
                                    }
                                }
                                for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                                {
                                    if (CheckBoxList_array01.Items[i].Selected)
                                    {
                                        CheckBoxList_array01.Items[i].Selected = false;
                                    }
                                }
                                for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                                {
                                    if (CheckBoxList_array02.Items[i].Selected)
                                    {
                                        CheckBoxList_array02.Items[i].Selected = false;
                                    }
                                }
                                for (int i = 0; i < CheckBoxList_Services.Items.Count; i++)
                                {
                                    if (CheckBoxList_Services.Items[i].Selected)
                                    {
                                        CheckBoxList_Services.Items[i].Selected = false;
                                    }
                                }
                                ResultBox.Text += "DateTime end: " + DateTime.Now.ToString() + "\r\n";
                                ResultBox.Text += "+++++ END +++++";
                                File.AppendAllText(stringPathLog, ResultBox.Text);
                            }
                            else { ResultBox.Text = "Please, select service."; }
                        }
                        else { ResultBox.Text = "You cannot choose servers from both array."; }
                    }
                    else { ResultBox.Text = "Please, select Mailbox server."; }
                }
                else { ResultBox.Text = "Wrong Input Code!"; }
            }
        }
        protected void Restart_click(object sender, EventArgs e)
        {
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                if (TextBox1.Text == input)
                {
                    ResultBox.Text = string.Empty;
                    var builder = new StringBuilder();
                    List<string> arrayMailboxServers = new List<string>();
                    //List<string> arrayMailboxServers_array02 = new List<string>();
                    List<string> arrayServices = new List<string>();
                    for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                    {
                        if (CheckBoxList_array01.Items[i].Selected)
                        {
                            arrayMailboxServers.Add(CheckBoxList_array01.Items[i].Text);
                        }
                    }
                    for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                    {
                        if (CheckBoxList_array02.Items[i].Selected)
                        {
                            arrayMailboxServers.Add(CheckBoxList_array02.Items[i].Text);
                        }
                    }
                    for (int i = 0; i < CheckBoxList_Services.Items.Count; i++)
                    {
                        if (CheckBoxList_Services.Items[i].Selected)
                        {
                            string stringReplaceService = Regex.Replace(CheckBoxList_Services.Items[i].Text, @".*[(]", "");
                            string stringReplaceService1 = Regex.Replace(stringReplaceService, @"[)]", "");
                            arrayServices.Add(stringReplaceService1);
                        }
                    }
                    if (arrayMailboxServers.Count > 0)
                    {
                        string stringMailboxServers = string.Join(" ", arrayMailboxServers);
                        boolServers_array01 = Regex.IsMatch(stringMailboxServers, @"nr", RegexOptions.IgnoreCase);
                        boolServers_array02 = Regex.IsMatch(stringMailboxServers, @"nt", RegexOptions.IgnoreCase);
                        if ((boolServers_array01 & !boolServers_array02) || (!boolServers_array01 & boolServers_array02))
                        {
                            if (arrayServices.Count > 0)
                            {
                                ResultBox.Text = "+++++ START +++++" + "\r\n";
                                ResultBox.Text += "DateTime start: " + DateTime.Now.ToString() + "\r\n";
                                ResultBox.Text += "Username: " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + "\r\n";
                                foreach (object objectServer in arrayMailboxServers)
                                {
                                    string stringMailboxServer = (string)objectServer;
                                    foreach (object objectService in arrayServices)
                                    {
                                        string stringService = (string)objectService;
                                        PingReply reply = pinger.Send(stringMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        if (pingable)
                                        {
                                            script = @"
                                            $password = ConvertTo-SecureString 'Password' -AsPlainText -Force
                                            $cred = New-Object System.Management.Automation.PSCredential('Username', $password)
                                            Invoke-Command -ComputerName '" + stringMailboxServer + @"' -ScriptBlock {Restart-Service '" + stringService + @"'} -Credential $cred
                                            if($?){
                                                Write-Output 'Successful RESTARTED service - " + stringService + @" on server - " + stringMailboxServer + @"!'
                                                Invoke-Command -ComputerName '" + stringMailboxServer + @"' -ScriptBlock {Get-Service -Name '" + stringService + @"'} -Credential $cred | select PSComputerName,name,status
                                            }else{
                                                Write-Output 'ERROR'; $error; break
                                            }
                                            ";
                                            using (ps = PowerShell.Create())
                                            {
                                                PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                                ps.AddScript(script);
                                                IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                                ps.EndInvoke(invocation);
                                                ps.Stop();
                                                foreach (PSObject outp in output)
                                                {
                                                    if (!outp.BaseObject.ToString().Contains("tmp"))
                                                    {
                                                        ResultBox.Text += outp.ToString() + "\r\n";
                                                    }
                                                }
                                            }
                                            ResultBox.Text += builder.ToString();
                                        }
                                        else { ResultBox.Text += "\r\n" + "No ping -> " + stringMailboxServer; }
                                    }
                                }
                                for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                                {
                                    if (CheckBoxList_array01.Items[i].Selected)
                                    {
                                        CheckBoxList_array01.Items[i].Selected = false;
                                    }
                                }
                                for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                                {
                                    if (CheckBoxList_array02.Items[i].Selected)
                                    {
                                        CheckBoxList_array02.Items[i].Selected = false;
                                    }
                                }
                                for (int i = 0; i < CheckBoxList_Services.Items.Count; i++)
                                {
                                    if (CheckBoxList_Services.Items[i].Selected)
                                    {
                                        CheckBoxList_Services.Items[i].Selected = false;
                                    }
                                }
                                ResultBox.Text += "DateTime end: " + DateTime.Now.ToString() + "\r\n";
                                ResultBox.Text += "+++++ END +++++";
                                File.AppendAllText(stringPathLog, ResultBox.Text);
                            }
                            else { ResultBox.Text = "Please, select service."; }
                        }
                        else { ResultBox.Text = "You cannot choose servers from both array."; }
                    }
                    else { ResultBox.Text = "Please, select Mailbox server."; }
                }
                else { ResultBox.Text = "Wrong Input Code!"; }
            }
        }
        protected void CheckService_click(object sender, EventArgs e)
        {
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                ResultBox.Text = string.Empty;
                var builder = new StringBuilder();
                List<string> arrayMailboxServers = new List<string>();
                List<string> arrayServices = new List<string>();
                for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                {
                    if (CheckBoxList_array01.Items[i].Selected)
                    {
                        arrayMailboxServers.Add(CheckBoxList_array01.Items[i].Text);
                    }
                }
                for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                {
                    if (CheckBoxList_array02.Items[i].Selected)
                    {
                        arrayMailboxServers.Add(CheckBoxList_array02.Items[i].Text);
                    }
                }
                for (int i = 0; i < CheckBoxList_Services.Items.Count; i++)
                {
                    if (CheckBoxList_Services.Items[i].Selected)
                    {
                        string stringReplaceService = Regex.Replace(CheckBoxList_Services.Items[i].Text, @".*[(]", "");
                        string stringReplaceService1 = Regex.Replace(stringReplaceService, @"[)]", "");
                        arrayServices.Add(stringReplaceService1);
                    }
                }
                ResultBox.Text = "+++++ START +++++" + "\r\n";
                            ResultBox.Text += "DateTime start: " + DateTime.Now.ToString() + "\r\n";
                            ResultBox.Text += "Username: " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + "\r\n";
                            foreach (object objectServer in arrayMailboxServers)
                            {
                                string stringMailboxServer = (string)objectServer;
                                foreach (object objectService in arrayServices)
                                {
                                    string stringService = (string)objectService;
                                    PingReply reply = pinger.Send(stringMailboxServer);
                                    pingable = reply.Status == IPStatus.Success;
                                    if (pingable)
                                    {
                                        script = @"
                                        $password = ConvertTo-SecureString 'Password' -AsPlainText -Force
                                        $cred = New-Object System.Management.Automation.PSCredential('Username', $password)
                                        Invoke-Command -ComputerName '" + stringMailboxServer + @"' -ScriptBlock {Get-Service -Name '" + stringService + @"'} -Credential $cred | select PSComputerName,name,status
                                        if(!$?){Write-Output 'ERROR'; $error; break}
                                        ";
                                        using (ps = PowerShell.Create())
                                        {
                                            PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                            ps.AddScript(script);
                                            IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                            ps.EndInvoke(invocation);
                                            ps.Stop();
                                            foreach (PSObject outp in output)
                                            {
                                                if (!outp.BaseObject.ToString().Contains("tmp"))
                                                {
                                                    ResultBox.Text += outp.ToString() + "\r\n";
                                                }
                                            }
                                        }
                                        ResultBox.Text += builder.ToString();
                                    }
                                    else { ResultBox.Text += "\r\n" + "No ping -> " + stringMailboxServer; }
                                }
                            }
                            ResultBox.Text += "DateTime end: " + DateTime.Now.ToString() + "\r\n";
                            ResultBox.Text += "+++++ END +++++";
            }
        }
    }
}