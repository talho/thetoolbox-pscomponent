using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.EnterpriseServices;
using System.Security;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Xml.Serialization;

namespace PowerShellComponent
{
    public class ManagementCommands : System.EnterpriseServices.ServicedComponent
    {
        public string EnableMailbox(string identity, string alias)
        {
            String ErrorText = "";
            String ReturnSet = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;

            // Load Exchange PowerShell snap-in.
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;

            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Enable-Mailbox");
                        thisPipeline.Commands[0].Parameters.Add("Identity", @identity);
                        thisPipeline.Commands[0].Parameters.Add("Alias", @alias);
                        thisPipeline.Commands[0].Parameters.Add("Database", @"mail2007.thetoolbox.com\First Storage Group\Mailbox Database");
                        thisPipeline.Commands[0].Parameters.Add("DomainController", "adtest2003.thetoolbox.com");
                        thisPipeline.Invoke();
                        try
                        {
                            ReturnSet = GetUser(identity);
                        }
                        catch (Exception ex)
                        {
                            ErrorText = "Error: " + ex.ToString();
                            return ErrorText;
                        }

                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling Enable-Mailbox.");
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }

                            ErrorText = ErrorText + "Error: " + pipelineError.ToString();
                        }
                    }
                }

                finally
                {
                    thisRunspace.Close();
                }
            }
            return ReturnSet;
        }



        public string NewADUser(string name, string externalEmailAddress, string password, string upn, string ou, string identity)
        {
            String ErrorText = "";
            String ReturnSet = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;

            // Load Exchange PowerShell snap-in.
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;

            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("New-MailUser");
                        thisPipeline.Commands[0].Parameters.Add("Name", @name);
                        thisPipeline.Commands[0].Parameters.Add("ExternalEmailAddress", @externalEmailAddress);
                        thisPipeline.Commands[0].Parameters.Add("Password", @password);
                        thisPipeline.Commands[0].Parameters.Add("UserPrincipalName", @upn);
                        thisPipeline.Commands[0].Parameters.Add("OrganizationalUnit", @ou);
                        thisPipeline.Commands[0].Parameters.Add("Database", @"mail2007.thetoolbox.com\First Storage Group\Mailbox Database");
                        thisPipeline.Commands[0].Parameters.Add("DomainController", "adtest2003.thetoolbox.com");
                        thisPipeline.Invoke();
                        try
                        {
                            ReturnSet = GetUser(identity);
                        }
                        catch (Exception ex)
                        {
                            ErrorText = "Error: " + ex.ToString();
                            return ErrorText;
                        }

                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling New-MailUser.");
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }

                            ErrorText = ErrorText + "Error: " + pipelineError.ToString();
                        }
                    }
                }
                finally
                {
                    thisRunspace.Close();
                }
            }
            return "";
        }
        public bool DeleteUser(string identity)
        {
            String ReturnSet = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;

            // Load Exchange PowerShell snap-in.
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;

            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Remove-Mailbox");
                        thisPipeline.Commands[0].Parameters.Add("Identity", identity);
                        thisPipeline.Commands[0].Parameters.Add("Confirm", false);
                        thisPipeline.Commands[0].Parameters.Add("DomainController", "adtest2003.thetoolbox.com");

                        try
                        {
                            thisPipeline.Invoke();
                            ReturnSet = "True";
                        }
                        catch (Exception ex)
                        {
                            ReturnSet = "Error: " + ex.ToString();
                        }

                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling Remove-Mailbox.");
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }

                            ReturnSet = ReturnSet + "Error: " + pipelineError.ToString();
                        }
                    }
                }

                finally
                {
                    thisRunspace.Close();
                }
            }
            if (ReturnSet == "True")
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public string GetUser(string identity)
        {
            String ErrorText = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            ExchangeUser user = null;
            List<ExchangeUser> users = new List<ExchangeUser>();

            // Load Exchange PowerShell snap-in.
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;

            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Get-User");
                        if(identity.Length > 0) thisPipeline.Commands[0].Parameters.Add("Identity", @identity);

                        try
                        {
                            foreach (PSObject result in thisPipeline.Invoke())
                            {
                                user = new ExchangeUser();
                                foreach (PSMemberInfo member in result.Members)
                                {
                                    switch (member.Name)
                                    {
                                        case "FirstName":
                                            user.givenName = member.Value.ToString().Trim();
                                            break;
                                        case "LastName":
                                            user.sn = member.Value.ToString().Trim();
                                            break;
                                        case "DistinguishedName":
                                            user.dn = member.Value.ToString().Trim();
                                            break;
                                        case "Name":
                                            user.cn = member.Value.ToString().Trim();
                                            break;
                                        case "UserPrincipalName":
                                            user.upn = member.Value.ToString().Trim();
                                            break;
                                    }
                                }
                                if (user.upn.Length > 0)
                                {
                                    
                                        using (Pipeline newPipeline = thisRunspace.CreatePipeline())
                                        {
                                            newPipeline.Commands.Add("Get-Mailbox");
                                            newPipeline.Commands[0].Parameters.Add("Identity", @user.upn);
                                            foreach (PSObject result2 in newPipeline.Invoke())
                                            {
                                                user.mailboxEnabled = (bool)result2.Members["IsValid"].Value;
                                            }
                                        }
                                    users.Add(user);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            ErrorText = "Error: " + ex.ToString();
                            return ErrorText;
                        }

                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling Enable-Mailbox.");
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }

                            ErrorText = ErrorText + "Error: " + pipelineError.ToString();
                        }
                    }
                }

                finally
                {
                    thisRunspace.Close();
                }
            }
            if (users.Count == 0)
            {
                return null;
            }
            else if (users.Count > 1)
            {

                XmlSerializer serializer = new XmlSerializer(typeof(List<ExchangeUser>));
                StringWriter textWriter = new StringWriter();
                serializer.Serialize(textWriter, users);
                textWriter.Close();
                return textWriter.ToString();
            }
            else
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ExchangeUser));
                StringWriter textWriter = new StringWriter();
                if (users.Count == 0)
                    serializer.Serialize(textWriter, new ExchangeUser());
                else
                    serializer.Serialize(textWriter, users[0]);
                textWriter.Close();
                return textWriter.ToString();
            }
        }

        public string GetIdentity()
        {
            AppDomain.CurrentDomain.SetPrincipalPolicy(System.Security.Principal.PrincipalPolicy.WindowsPrincipal);
            System.Security.Principal.WindowsPrincipal user = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
            return user.Identity.Name;
        }
    }
}