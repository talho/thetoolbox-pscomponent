﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.EnterpriseServices;
using System.Security;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Linq;
using System.Xml.Serialization;
using ToolBoxUtility;
using System.Configuration;

// Scope PowerShellComponent
namespace PowerShellComponent
{
    // Class ManagementCommands
    public class ManagementCommands : System.EnterpriseServices.ServicedComponent
    {
        private KeyValueConfigurationCollection appSettings = null;

        private KeyValueConfigurationCollection AppSettings
        {
            get
            {
                if (appSettings == null)
                {
                    appSettings = ConfigurationManager.OpenExeConfiguration(System.Reflection.Assembly.GetExecutingAssembly().Location).AppSettings.Settings;
                }
                return appSettings;
            }
        }

        #region Users and Mailboxes

        // EnableMailbox()
        // desc: Method uses PowerShellSnapIn "Microsoft.Exchange.Management.PowerShell.Admin" to enable mailbox on Exchange using CMDLET Enable-Mailbox
        // params: string identity - User login name
        //         string alias    - User login name
        // method: public
        // return: string, ExchangeUser XML serialized
        public string EnableMailbox(string identity, string alias)
        {
            String ErrorText             = "";
            String ReturnSet             = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            // Load Exchange PowerShell snap-in.
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;
            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config)){
                try{
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline()){
                        thisPipeline.Commands.Add("Enable-Mailbox");
                        thisPipeline.Commands[0].Parameters.Add("Identity", @identity);
                        thisPipeline.Commands[0].Parameters.Add("Alias", @alias);
                        thisPipeline.Commands[0].Parameters.Add("Database", AppSettings["database"].Value);
                        thisPipeline.Commands[0].Parameters.Add("DomainController", AppSettings["domainController"].Value);
                        thisPipeline.Invoke();
                        try{
                            ReturnSet = GetUser(identity);
                        }catch (Exception ex){
                            ErrorText = "Error: " + ex.ToString();
                            return ErrorText;
                        }
                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0){
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling Enable-Mailbox.");
                            foreach (object item in thisPipeline.Error.ReadToEnd()){
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            ErrorText = ErrorText + "Error: " + pipelineError.ToString();
                        }
                    }
                }
                finally{
                    thisRunspace.Close();
                }
            }
            return ReturnSet;
        }

        // NewADUser()
        // desc: Method loads RunSpace, imports ActiveDirectory module into PS session, creates new user in AD server only using CMDLET New-ADUser
        // params: Dictionary<string, string> attributes - Dictionary object, contains attributes for creating a new user
        // method: public
        // return: string, ExchangeUser XML serialized
        public string NewADUser(Dictionary<string, string> attributes)
        {
            String ErrorText             = "";
            String ReturnSet             = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config)){
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Import-Module");
                        thisPipeline.Commands[0].Parameters.Add("Name", "ActiveDirectory");
                        thisPipeline.Invoke();
                    }
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("New-ADUser");
                        thisPipeline.Commands[0].Parameters.Add("Name", @attributes["name"]);
                        thisPipeline.Commands[0].Parameters.Add("DisplayName", @attributes["displayName"]);
                        thisPipeline.Commands[0].Parameters.Add("GivenName", @attributes["givenName"]);
                        thisPipeline.Commands[0].Parameters.Add("Surname", @attributes["sn"]);
                        SecureString secureString = new SecureString();
                        foreach (char c in @attributes["password"])
                            secureString.AppendChar(c);
                        secureString.MakeReadOnly();
                        thisPipeline.Commands[0].Parameters.Add("AccountPassword", secureString);
                        thisPipeline.Commands[0].Parameters.Add("UserPrincipalName", @attributes["upn"]);
                        thisPipeline.Commands[0].Parameters.Add("Path", @attributes["dn"]);
                        thisPipeline.Commands[0].Parameters.Add("SamAccountName", @attributes["samAccountName"]);
                        thisPipeline.Commands[0].Parameters.Add("PasswordNeverExpires", Int32.Parse(@attributes["pwdExpires"]));
                        thisPipeline.Commands[0].Parameters.Add("ChangePasswordAtLogon", Int32.Parse(@attributes["changePwd"]));
                        thisPipeline.Commands[0].Parameters.Add("Enabled", Int32.Parse(@attributes["acctDisabled"]));
                        thisPipeline.Invoke();
                        try{
                            //ReturnSet = GetUser(attributes["alias"].Replace("-vpn", ""));
                            ReturnSet = GetUser(@attributes["upn"]);
                        }catch (Exception ex){
                            ErrorText = "Error: " + ex.ToString();
                            return ErrorText;
                        }
                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0){
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling New-ADUser.");
                            foreach (object item in thisPipeline.Error.ReadToEnd()){
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            ErrorText = ErrorText + "Error: " + pipelineError.ToString();
                            return ErrorText;
                        }
                    }
                }catch (Exception ex){
                    ErrorText = "Error: " + ex.ToString();
                    return ErrorText;
                }finally{
                    thisRunspace.Close();
                }
            }
            return ReturnSet;
        }

        // NewExchangeUser()
        // desc: Method uses PowerShellSnapIn "Microsoft.Exchange.Management.PowerShell.Admin" to create a new user on Exchange using CMDLET New-Mailbox
        // params: Dictionary<string, string> attributes - Dictionary Object with attributes for creating new user
        // method: public
        // return: string, ExchangeUser XML serialized
        public string NewExchangeUser(Dictionary<string, string> attributes)
        {
            String ErrorText             = "";
            String ReturnSet             = "";
            RunspaceConfiguration config = null;
            PSSnapInException warning;
            Runspace thisRunspace = null;
            try{
                config = RunspaceConfiguration.Create();
                // Load Exchange PowerShell snap-in.
                config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
                if (warning != null) throw warning;

                using (thisRunspace = RunspaceFactory.CreateRunspace(config)){
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline()){
                        thisPipeline.Commands.Add("New-Mailbox");
                        thisPipeline.Commands[0].Parameters.Add("Name", @attributes["name"]);
                        thisPipeline.Commands[0].Parameters.Add("Alias", @attributes["alias"]);
                        thisPipeline.Commands[0].Parameters.Add("FirstName", @attributes["givenName"]);
                        thisPipeline.Commands[0].Parameters.Add("LastName", @attributes["sn"]);
                        thisPipeline.Commands[0].Parameters.Add("DisplayName", @attributes["displayName"]);
                        SecureString secureString = new SecureString();
                        foreach(char c in @attributes["password"])
                                secureString.AppendChar(c);
                        secureString.MakeReadOnly();
                        thisPipeline.Commands[0].Parameters.Add("Password", secureString);
                        thisPipeline.Commands[0].Parameters.Add("UserPrincipalName", @attributes["upn"]);
                        thisPipeline.Commands[0].Parameters.Add("OrganizationalUnit", @attributes["ou"]);
                        thisPipeline.Commands[0].Parameters.Add("ResetPasswordOnNextLogon", Int32.Parse(@attributes["changePwd"]));
                        thisPipeline.Commands[0].Parameters.Add("Database", AppSettings["database"].Value);
                        thisPipeline.Invoke();
                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling New-MailUser.");
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            throw new Exception(ErrorText + "Error: " + pipelineError.ToString());
                        }
                        else
                        {
                            try
                            {
                                ReturnSet = GetUser(attributes["alias"]);
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Error: " + ex.ToString());
                            }
                        }
                    }

                    // After Mailbox is created, set certain attributes
                    // rewrite CustomAttribute1 to be OU
                    // rewrite msExchQueryBaseDN
                    // rewrite msExchUseOAB by setting OfflineAddressBook attribute
                    using (Pipeline pipey = thisRunspace.CreatePipeline())
                    {
                        pipey.Commands.Add("Set-Mailbox");
                        pipey.Commands[0].Parameters.Add("Identity", @attributes["alias"]);
                        pipey.Commands[0].Parameters.Add("DomainController", AppSettings["domainController"].Value);
                        pipey.Commands[0].Parameters.Add("CustomAttribute1", @attributes["ou"]);
                        pipey.Commands[0].Parameters.Add("OfflineAddressBook", @attributes["useOAB"]);
                        pipey.Invoke();
                        using (Pipeline pipeCommand = thisRunspace.CreatePipeline())
                        {
                            String setQueryBaseDNCommand = "Get-Mailbox -Identity " + @attributes["alias"]
                                                        + " | foreach{$dn='LDAP://'+$_.distinguishedName;"
                                                        + "$obj=[ADSI]$dn;$obj.msExchQueryBaseDN='" + @attributes["dn"] + "';$obj.setInfo();}";
                            pipeCommand.Commands.AddScript(setQueryBaseDNCommand);
                            pipeCommand.Invoke();
                        }
                        
                    }
                }
            }catch (Exception ex){
                ErrorText = "Error: " + ex.ToString();
                return ErrorText;
            }finally{
                thisRunspace.Close();
            }
            return ReturnSet;
        }

        // AddGroupMember()
        // desc: Method uses Powershell Cmdlet Add-ADGroupMember
        // params: Dictionary<string, string> attributes - Dictionary Object with attributes for adding member to group
        // method: public
        // return: bool
        public string AddGroupMember(Dictionary<string, string> attributes)
        {
            String ErrorText             = "";
            String ReturnSet             = "";
            RunspaceConfiguration config = null;
            PSSnapInException warning;
            Runspace thisRunspace = null;
            try
            {
                config = RunspaceConfiguration.Create();
                // Load Exchange PowerShell snap-in.
                config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
                if (warning != null) throw warning;

                using (thisRunspace = RunspaceFactory.CreateRunspace(config))
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Import-Module");
                        thisPipeline.Commands[0].Parameters.Add("Name", "ActiveDirectory");
                        thisPipeline.Invoke();
                    }
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Add-ADGroupMember");
                        thisPipeline.Commands[0].Parameters.Add("Identity", @attributes["securityGroup"]);
                        thisPipeline.Commands[0].Parameters.Add("Member", @attributes["alias"]);
                        thisPipeline.Invoke();
                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling Add-ADGroupMember.");
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            throw new Exception(ErrorText + "Error: " + pipelineError.ToString());
                        }
                        else
                        {
                            try
                            {
                                ReturnSet = GetUser(attributes["alias"]);
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Error: " + ex.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorText = "Error: " + ex.ToString();
                return ErrorText;
            }
            finally
            {
                thisRunspace.Close();
            }
            return ReturnSet;
        }

        // DeleteUser()
        // desc: Method uses PowerShellSnapIn "Microsoft.Exchange.Management.PowerShell.Admin" to delete user using CMDLET Remove-Mailbox
        // params: string identity - User login name
        // method: public
        // return: bool
        public bool DeleteUser(string identity)
        {
            String ReturnSet             = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            // Load Exchange PowerShell snap-in.
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;
            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config)){
                try{
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline()){
                        thisPipeline.Commands.Add("Remove-Mailbox");
                        thisPipeline.Commands[0].Parameters.Add("Identity", identity);
                        thisPipeline.Commands[0].Parameters.Add("Confirm", false);
                        thisPipeline.Commands[0].Parameters.Add("DomainController", AppSettings["domainController"].Value);
                        try{
                            thisPipeline.Invoke();
                            ReturnSet = "True";
                        }catch (Exception ex){
                            ReturnSet = "Error: " + ex.ToString();
                        }
                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0){
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling Remove-Mailbox.");
                            foreach (object item in thisPipeline.Error.ReadToEnd()){
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            ReturnSet = ReturnSet + "Error: " + pipelineError.ToString();
                        }
                    }
                }finally{
                    thisRunspace.Close();
                }
            }if (ReturnSet == "True"){
                return true;
            }else{
                return false;
            }
        }

        public bool DeleteADUser(string identity)
        {
            String ReturnSet = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Import-Module");
                        thisPipeline.Commands[0].Parameters.Add("Name", "ActiveDirectory");
                        thisPipeline.Invoke();
                    }
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Remove-ADUser");
                        thisPipeline.Commands[0].Parameters.Add("Identity", identity);
                        thisPipeline.Commands[0].Parameters.Add("Confirm", false);
                        try
                        {
                            thisPipeline.Invoke();
                            ReturnSet = "True";
                        }
                        catch (Exception ex)
                        {
                            ReturnSet = "Error: " + ex.ToString();
                            throw new Exception(ReturnSet);
                        }
                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling Remove-ADUser.");
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
            } if (ReturnSet == "True")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        // ChangePassword()
        // desc: Method loads RunSpace, imports ActiveDirectory module into PS session, uses Set-ADAccountPassword CMDLET to reset user password
        // params: Dictionary<string, string> attributes - Dictionary Object with attributes for changing a user password
        // method: public
        // return: bool
        public bool ChangePassword(string identity, string password)
        {
            String ReturnSet = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            
            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Import-Module");
                        thisPipeline.Commands[0].Parameters.Add("Name", "ActiveDirectory");
                        thisPipeline.Invoke();
                    }
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Set-ADAccountPassword");
                        thisPipeline.Commands[0].Parameters.Add("Identity", @identity);
                        thisPipeline.Commands[0].Parameters.Add("Reset", true);
                        SecureString secureString = new SecureString();
                        foreach (char c in @password)
                            secureString.AppendChar(c);
                        secureString.MakeReadOnly();
                        thisPipeline.Commands[0].Parameters.Add("NewPassword", secureString);
                        try{
                            thisPipeline.Invoke();
                            ReturnSet = "True";
                        }catch (Exception ex){
                            ReturnSet = "Error: " + ex.ToString();
                        }
                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0){
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling Set-ADAccountPassword.");
                            foreach (object item in thisPipeline.Error.ReadToEnd()){
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            ReturnSet = ReturnSet + "Error: " + pipelineError.ToString();
                        }
                    }
                }finally{
                    thisRunspace.Close();
                }
              
            }
            if (ReturnSet == "True"){
                return true;
            }else{
                return false;
            }

        }

        /// <summary>
        /// Method uses PowerShellSnapIn "Microsoft.Exchange.Management.PowerShell.Admin" to get user mailbox using CMDLET Get-Mailbox
        /// </summary>
        /// <param name="identity">User login name</param>
        /// <param name="current_page">Current Page to return</param>
        /// <param name="per_page">Entries to return per page</param>
        /// <param name="vpn_only">Return only VPN users, default is false</param>
        /// <returns>String, ExchangeUser XML serialized</returns>
        public string GetUser(string identity, int current_page = 1, int per_page = 10, bool vpn_only = false, string ou = "")
        {
            List<ExchangeUser> users = new List<ExchangeUser>();
            ExchangeUserShorter shorty = new ExchangeUserShorter() { CurrentPage = current_page, PerPage = per_page };
            int total_entries;

            try
            {
                users = GetUsers(out total_entries, identity:identity, displayName:"", current_page:current_page, per_page:per_page, vpn_only:vpn_only, ou:ou);
                shorty.users = users;
                shorty.TotalEntries = total_entries;
            }
            catch (Exception e)
            {
                return e.Message;
            }

            if (users.Count == 0){
                return XmlSerializationHelper.Serialize(shorty);
            }else if (identity.Trim() == ""){
                return XmlSerializationHelper.Serialize(shorty);// +"THEWORLDSLARGESTSEPERATOR" + total_entries.ToString();
            }else{
                return XmlSerializationHelper.Serialize(users[0]);
            }
        }

        /// <summary>
        /// Gets a list of users based on the provided identity, paged if provided.
        /// </summary>
        /// <param name="total_entries">output parameter of the total number of users found</param>
        /// <param name="identity">the alias of the user we want to find. If it's blank, the method will find all users</param>
        /// <param name="displayName">The display name of the user. Not used right now, may be later</param>
        /// <param name="current_page">The current page, for paging, that we are on</param>
        /// <param name="per_page">The number of users to display per page</param>
        /// <param name="vpn_only">Return only VPN users, default is false</param>
        /// <returns>If identity is blank, returns a list with all users. If identity is not blank, returns all users</returns>
        private List<ExchangeUser> GetUsers(out int total_entries, string identity = "", string displayName = "", int current_page = 0, int per_page = 0, bool vpn_only = false, string ou = "")
        {
            String ErrorText = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            ExchangeUser user = null;
            List<ExchangeUser> users = new List<ExchangeUser>();
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
                        if (identity.IndexOf("-vpn") != -1 || vpn_only) thisPipeline.Commands.Add("Get-User");
                        else thisPipeline.Commands.Add("Get-Mailbox");
                        if (ou != "") thisPipeline.Commands[0].Parameters.Add("OrganizationalUnit", @ou);
                        if (identity != "") thisPipeline.Commands[0].Parameters.Add("Identity", @identity);
                        else if(vpn_only) thisPipeline.Commands[0].Parameters.Add("Filter", "SamAccountName -like '*vpn*'");
                        if (displayName != "") thisPipeline.Commands[0].Parameters.Add("Anr", @displayName);
                        thisPipeline.Commands[0].Parameters.Add("SortBy", "DisplayName");
                        thisPipeline.Commands[0].Parameters.Add("ResultSize", "Unlimited");
                        try
                        {
                            List<PSObject> original_results = thisPipeline.Invoke().ToList();
                            IEnumerable<PSObject> results = null;

                            //We need to filter the results further by OU, if ou is set in order to filter out possible child ou's from result list
                            if(ou != "" && !vpn_only)
                                original_results = original_results.Where(x => x.Members["OrganizationalUnit"].Value.ToString() == ou).ToList();
                            else if (ou != "" && vpn_only)
                                original_results = original_results.Where(x => x.Members["Identity"].Value.ToString() == (ou + "/VPN/" + x.Members["Name"].Value.ToString())).ToList();
                            
                            total_entries = original_results.Count;
                            
                            if (current_page == 0 && per_page == 0)
                                results = original_results;
                            else if (current_page < 2)
                                results = original_results.Take<PSObject>(per_page);
                            else
                                results = original_results.Skip<PSObject>((current_page - 1) * per_page).Take<PSObject>(per_page);

                            foreach (PSObject result in results)
                            {
                                user = ReadUserInformation(result);

                                if (user.upn != "" && !vpn_only)
                                {
                                    using (Pipeline newPipeline = thisRunspace.CreatePipeline())
                                    {
                                        //if (user.identity.IndexOf("-vpn") == -1)
                                        //{
                                            string vpn_identity = user.upn.Replace("@", "-vpn@");
                                            //vpn_identity = user.upn;
                                            newPipeline.Commands.Add("Get-User");
                                            newPipeline.Commands[0].Parameters.Add("Identity", @vpn_identity);
                                            foreach (PSObject result2 in newPipeline.Invoke())
                                            {
                                                user.has_vpn = (((string)result2.Members["UserPrincipalName"].Value).Length > 0);
                                            }
                                        //}
                                    }
                                }

                                users.Add(user);
                            }

                        }
                        catch (Exception ex)
                        {
                            ErrorText = "Error: " + ex.ToString();
                            throw new Exception(ErrorText);
                        }
                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling Get-Mailbox.");
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            ErrorText = ErrorText + "Error: " + pipelineError.ToString();
                            throw new Exception(ErrorText);
                        }
                    }
                }
                finally
                {
                    thisRunspace.Close();
                }
            }

            return users;
        }

        /// <summary>
        /// Reads a user's properties from the PSObject that is provided
        /// </summary>
        /// <param name="result">A powershell object with member elements</param>
        /// <returns>An exchange user with properties filled out</returns>
        private ExchangeUser ReadUserInformation(PSObject result)
        {
            ExchangeUser user = new ExchangeUser();
            foreach (PSMemberInfo member in result.Members)
            {
                switch (member.Name)
                {
                    case "Alias":
                        user.alias = member.Value.ToString().Trim();
                        break;
                    case "DistinguishedName":
                        user.dn = member.Value.ToString().Trim();
                        break;
                    case "DisplayName":
                        user.cn = member.Value.ToString().Trim();
                        break;
                    case "UserPrincipalName":
                        user.upn = member.Value.ToString().Trim();
                        break;
                    case "SamAccountName":
                        user.login = member.Value.ToString().Trim();
                        break;
                    case "OrganizationalUnit":
                        user.ou = member.Value.ToString().Trim();
                        //user.ou = user.ou.Substring(user.ou.IndexOf('/') + 1);
                        break;
                    case "PrimarySmtpAddress":
                    case "WindowsEmailAddress":
                        user.email = member.Value.ToString().Trim();
                        break;
                    case "IsValid":
                        user.mailboxEnabled = (bool)member.Value;
                        break;
                    case "RecipientType":
                        user.type = member.Value.ToString().Trim();
                        break;
                }
            }
            return user;
        }
        
        #endregion

        #region Distribution Groups and management

        // CreateDistributionGroup()
        // desc: Method creates a new distribution list
        // params: string group_name - Name of new disitribution list
        // method: public
        // return: bool
        public string CreateDistributionGroup(string group_name, string ou, string auth_enabled)
        {
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;

            DistributionGroup group = new DistributionGroup() { error = "" };

            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        try
                        {
                            thisPipeline.Commands.Add("New-DistributionGroup");
                            thisPipeline.Commands[0].Parameters.Add("Name", @group_name);
                            thisPipeline.Commands[0].Parameters.Add("Type", @"Distribution");
                            thisPipeline.Commands[0].Parameters.Add("OrganizationalUnit", @ou);
                            thisPipeline.Commands[0].Parameters.Add("SamAccountName", @group_name);
                            thisPipeline.Commands[0].Parameters.Add("Alias", @group_name.Replace(" ", ""));
                            try
                            {
                                thisPipeline.Invoke();
                                DistributionGroupsShorter shorty = XmlSerializationHelper.Deserialize<DistributionGroupsShorter>(GetDistributionGroup(group_name, 0, 0, ""));
                                if (shorty.groups.Count > 0)
                                    group = shorty.groups[0];
                                else
                                    throw new Exception("Group creation failed somewhere, new group was not found");
                                
                                // rewrite CustomAttribute1 to be OU
                                using (Pipeline pipey = thisRunspace.CreatePipeline())
                                {
                                    pipey.Commands.Add("Set-DistributionGroup");
                                    pipey.Commands[0].Parameters.Add("Identity", group.Alias);
                                    pipey.Commands[0].Parameters.Add("DomainController", AppSettings["domainController"].Value);
                                    pipey.Commands[0].Parameters.Add("RequireSenderAuthenticationEnabled", Int32.Parse(@auth_enabled));
                                    pipey.Commands[0].Parameters.Add("CustomAttribute1", ou);
                                    pipey.Invoke();
                                }
                            }
                            catch (Exception ex)
                            {
                                group.error += " Error: " + ex.ToString();
                            }
                            // Check for errors in the pipeline and throw an exception if necessary.
                            if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                            {
                                StringBuilder pipelineError = new StringBuilder();
                                pipelineError.AppendFormat("Error calling New-DistributionGroup.");
                                foreach (object item in thisPipeline.Error.ReadToEnd())
                                {
                                    pipelineError.AppendFormat("{0}\n", item.ToString());
                                }
                                group.error += " Error: " + pipelineError.ToString();
                            }
                        }
                        catch (Exception ex)
                        {
                            group.error += " Error: " + ex.ToString();
                        }
                    }
                }
                finally
                {
                    thisRunspace.Close();
                }
            }

            return XmlSerializationHelper.Serialize(group);
        }

        // GetDistributionGroup()
        // desc: Method returns a list of Distribution Groups
        // params: sring identity - Name of Distribution group to return
        // method: public
        // return: string
        public string GetDistributionGroup(string identity, int current_page, int per_page, string ou)
        {
            String ErrorText = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            DistributionGroup group = null;
            List<DistributionGroup> groups = new List<DistributionGroup>();
            PSSnapInException warning;
            int total_entries = 0;
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
                        
                        thisPipeline.Commands.Add("Get-DistributionGroup");
                        if (identity != "") thisPipeline.Commands[0].Parameters.Add("Identity", @identity);
                        thisPipeline.Commands[0].Parameters.Add("SortBy", "DisplayName");
                        try
                        {                            
                            List<PSObject> original_results = thisPipeline.Invoke().ToList();
                            IEnumerable<PSObject> results = null;

                            if (ou != "")
                                original_results = original_results.Where(x => x.Members["OrganizationalUnit"].Value.ToString() == ou).ToList();
                            total_entries = original_results.Count;
                            if (current_page == 0 && per_page == 0)
                                results = original_results;
                            else if (current_page < 2)
                                results = original_results.Take<PSObject>(per_page);// + 1); // This one is working as you would expect, as opposed to users. not sure what was done there
                            else
                                results = original_results.Skip<PSObject>((current_page - 1) * per_page).Take<PSObject>(per_page);
                           foreach (PSObject result in results)
                           {
                               
                               group = new DistributionGroup();
                               
                               foreach (PSMemberInfo member in result.Members)
                               {
                                   switch (member.Name)
                                   {
                                       case "Alias":
                                           group.Alias = member.Value.ToString().Trim();
                                           break;
                                       case "Name":
                                           group.Name = member.Value.ToString().Trim();
                                           break;
                                       case "DisplayName":
                                           group.displayName = member.Value.ToString().Trim();
                                           break;
                                       case "GroupType":
                                           group.groupType = member.Value.ToString().Trim();
                                           break;
                                       case "PrimarySmtpAddress":
                                           group.primarySmtpAddress = member.Value.ToString();
                                           break;
                                       case "CustomAttribute11": // We're using attribute 11 to indicate if we have children or not. it will be blank or true most likely
                                           bool hasChildren;
                                           if (bool.TryParse(member.Value.ToString().Trim(), out hasChildren) && hasChildren)
                                               group.HasChildren = true;
                                           else
                                               group.HasChildren = false;
                                           break;
                                   }
                                     
                               }
                               if (group.displayName.Length > 0)
                               {
                                   group.users = new ExchangeUserMembers();
                                   if (identity != "")
                                       group.users.users = GetDistributionGroupMembers(group.Name);
                                   groups.Add(group);
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
                            pipelineError.AppendFormat("Error calling Get-DistributionGroup.");
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            ErrorText = ErrorText + "Error: " + pipelineError.ToString();
                            return ErrorText;
                        }
                          
                    }
                      
                }
                finally
                {
                    thisRunspace.Close();
                }
               
            }
            var shorty = new DistributionGroupsShorter() { PerPage = per_page, CurrentPage = current_page, 
                TotalEntries = total_entries, groups = groups };
            return XmlSerializationHelper.Serialize(shorty);                
        }

        /// <summary>
        /// Gets a list of users and contacts for the provided distribution group identity
        /// </summary>
        /// <param name="identity">The name/alias of a distribution group</param>
        /// <returns>A list of users that belong to the provided distribution group</returns>
        public List<ExchangeUser> GetDistributionGroupMembers(string identity)
        {
            String ErrorText = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            List<ExchangeUser> users = new List<ExchangeUser>();
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

                        thisPipeline.Commands.Add("Get-DistributionGroupMember");
                        thisPipeline.Commands[0].Parameters.Add("Identity", @identity);
                        try
                        {
                            Collection<PSObject> results = thisPipeline.Invoke();
                            foreach (PSObject result in results)
                            {
                                ExchangeUser user = ReadUserInformation(result);
                                
                                if(user.alias != "")
                                    users.Add(user);
                            }

                        }
                        catch (Exception ex)
                        {
                            ErrorText = "Error: " + ex.ToString();
                            throw new Exception(ErrorText);
                        }
                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling Get-DistributionGroup.");
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            ErrorText = ErrorText + "Error: " + pipelineError.ToString();
                            throw new Exception(ErrorText);
                        }

                    }

                }
                finally
                {
                    thisRunspace.Close();
                }
            }

            return users;
        }
        
        // AddToDistributionGroup()
        // desc: Method adds a member to an existing distribution group
        // params: string group_name - Name of new disitribution group
        //         string alias      - Alias of member to add
        // method: public
        // return: bool
        public bool AddToDistributionGroup(string group_name, string alias)
        {
            String ReturnSet = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;
            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        try
                        {
                            thisPipeline.Commands.Add("Add-DistributionGroupMember");
                            thisPipeline.Commands[0].Parameters.Add("identity", @group_name);
                            thisPipeline.Commands[0].Parameters.Add("member", @alias);
                            try
                            {
                                thisPipeline.Invoke();
                                ReturnSet = "True";
                            }
                            catch (Exception ex)
                            {
                                ReturnSet = "Error: " + ex.ToString();
                            }
                        }
                        catch (Exception ex)
                        {
                            ReturnSet = "Error: " + ex.ToString();
                        }
                    }
                }
                finally
                {
                    thisRunspace.Close();
                }
            }
            if (ReturnSet == "True")
                return true;
            else throw new Exception(ReturnSet);
                //return false;
        }

        /// <summary>
        /// Removes a single user from the provided distribution group with the provided alias
        /// </summary>
        /// <param name="group_name">The target distribution group</param>
        /// <param name="alias">The alias of the users to be removed</param>
        /// <returns>success</returns>
        public bool RemoveFromDistributionGroup(string group_name, string alias)
        {
            String ReturnSet = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;
            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        try
                        {
                            thisPipeline.Commands.Add("Remove-DistributionGroupMember");
                            thisPipeline.Commands[0].Parameters.Add("identity", @group_name);
                            thisPipeline.Commands[0].Parameters.Add("member", @alias);
                            thisPipeline.Commands[0].Parameters.Add("Confirm", false);
                            try
                            {
                                thisPipeline.Invoke();
                                ReturnSet = "True";
                            }
                            catch (Exception ex)
                            {
                                ReturnSet = "Error: " + ex.ToString();
                            }
                        }
                        catch (Exception ex)
                        {
                            ReturnSet = "Error: " + ex.ToString();
                        }
                    }
                }
                finally
                {
                    thisRunspace.Close();
                }
            }
            if (ReturnSet == "True")
                return true;
            else
                throw new Exception(ReturnSet);
        }

        /// <summary>
        /// The update distribution group method. Handles adding and removing group members and manages the group.HasChildren property.
        /// </summary>
        /// <remarks>Can be expanded to include other update features such as changing name, etc</remarks>
        /// <param name="distributionGroupXml">The XML representation of the group to be updated</param>
        /// <returns>The XML representation of the updated group</returns>
        public string UpdateDistributionGroup(string distributionGroupXml)
        {
            // We're going to update our users first. Othere things may follow, but users are the only one for now
            DistributionGroup group = XmlSerializationHelper.Deserialize<DistributionGroup>(distributionGroupXml);

            // get the current group member list
            List<ExchangeUser> currentMembers = GetDistributionGroupMembers(group.Name);
            
            // find the members that are new: those that don't exist in the currentMember list
            List<ExchangeUser> newUsers = group.users.users.Except(currentMembers, new LambdaComparer<ExchangeUser>((x, y) => (x.type != "MailContact" && y.type != "MailContact" && x.alias == y.alias) || (x.type == "MailContact" && y.type == "MailContact" && x.email == y.email) )).ToList();

            // find the members that are removed
            List<ExchangeUser> removedUsers = currentMembers.Except(group.users.users, new LambdaComparer<ExchangeUser>((x, y) => x.alias == y.alias)).ToList();

            // Add users using the existing methodology
            newUsers.FindAll(x => x.type == "MailContact").ForEach(x => CreateMailContact(ref x));
            newUsers.FindAll(x => x.error == "").ForEach(x => AddToDistributionGroup(group.Name, x.alias));

            // Do the remove here, if we're removing anything
            removedUsers.ForEach(x => RemoveFromDistributionGroup(group.Name, x.alias));

            group = XmlSerializationHelper.Deserialize<DistributionGroupsShorter>(GetDistributionGroup(group.Name, 0, 0, "")).groups[0];
            
            try
            {
                // We're going to reset the has children attribute, which we're storing on CustomAttribute11 here if it's not set properly
                if (!group.HasChildren && group.users.users.Count > 0)
                {
                    this.SetDistributionGroupChildren(group.Name, true);
                    group.HasChildren = true;
                }
                else if (group.HasChildren && group.users.users.Count == 0)
                {
                    this.SetDistributionGroupChildren(group.Name, false);
                    group.HasChildren = false;
                }
            }
            catch (Exception e)
            {
                group.error += e.Message; // If there's an error, we want to return that there was an error,
                                          // but there's no action to take on that error here.
            }
            
            // return the new group, though we won't use it.
            return XmlSerializationHelper.Serialize(group);
        }

        /// <summary>
        /// Sets on attribute CustomAttribute11 a string indicating if a group has children or not.
        /// </summary>
        /// <param name="groupName">The group that may or may not have children</param>
        /// <param name="hasChildren">If the group has children or not</param>
        public void SetDistributionGroupChildren(string groupName, bool hasChildren)
        {
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;
            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        try
                        {
                            thisPipeline.Commands.Add("Set-DistributionGroup");
                            thisPipeline.Commands[0].Parameters.Add("identity", @groupName);
                            thisPipeline.Commands[0].Parameters.Add("CustomAttribute11", @hasChildren);
                            try
                            {
                                thisPipeline.Invoke();
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
                finally
                {
                    thisRunspace.Close();
                }
            }
        }

        public void DeleteDistributionGroup(string distributionGroupXml)
        {
            DistributionGroup group = XmlSerializationHelper.Deserialize<DistributionGroup>(distributionGroupXml);
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;

            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        
                        thisPipeline.Commands.Add("Remove-DistributionGroup");
                        thisPipeline.Commands[0].Parameters.Add("Identity", group.Alias);
                        thisPipeline.Commands[0].Parameters.Add("Confirm", false);
                        thisPipeline.Invoke();

                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            throw new Exception(pipelineError.ToString());
                        }                        
                    }
                }
                finally
                {
                    thisRunspace.Close();
                }
            }
        }

        #endregion

        #region Mail Contacts

        /// <summary>
        /// Gets an existing mail contact
        /// </summary>
        /// <param name="alias">Alias of that contact</param>
        /// <returns>A mail contact or null if it doesn't exist</returns>
        protected ExchangeUser GetMailContact(string alias)
        {
            ExchangeUser contact = null;
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;

            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();

                    // first look for the contact to see if it already exists
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        
                        thisPipeline.Commands.Add("Get-MailContact");
                        thisPipeline.Commands[0].Parameters.Add("Identity", alias);
                        try
                        {
                            Collection<PSObject> results = thisPipeline.Invoke();

                            if (results.Count > 0)
                                contact = ReadUserInformation(results[0]);
                        }
                        catch { } // We don't really care about what went wrong, we're just going to say that the contact does not exist
                    }
                }
                finally
                {
                    thisRunspace.Close();
                }
            }

            return contact;
        }

        /// <summary>
        /// Takes an unknown contact and matches on a number of parameters for that contact in the order of:
        /// email, common name, alias
        /// </summary>
        /// <param name="contactMatch">the expanded contact that we want to try to match on</param>
        /// <returns></returns>
        protected ExchangeUser GetMailContact(ExchangeUser contactMatch)
        {
            ExchangeUser ret = GetMailContact(contactMatch.email);
            if (ret == null && contactMatch.cn != "")
            {
                ret = GetMailContact(contactMatch.cn);
            }
            if (ret == null && contactMatch.alias != "")
            {
                ret = GetMailContact(contactMatch.alias);
            }

            return ret;
        }

        public string GetSerializedMailContact(string contactMatch)
        {
            return XmlSerializationHelper.Serialize(GetMailContact(XmlSerializationHelper.Deserialize<ExchangeUser>(contactMatch)));
        }

        /// <summary>
        /// Function for creating a single contact. Calls CreateMailContact(ref ExchangeUser newContact, int limiter = 1)
        /// </summary>
        /// <param name="name">The contact name, assumed in the form of "First Last"</param>
        /// <param name="email">The contact's email address</param>
        /// <param name="ou">The organizational unit that this contact will belong to</param>
        /// <param name="alias">Optional alias for the user. If not provided, the alias will be set to whatever name was, minus the spaces</param>
        /// <returns>success</returns>
        public bool CreateMailContact(string name, string email, string ou, string alias = "")
        {
            ExchangeUser contact = new ExchangeUser() { cn = name, email = email, ou = ou, alias = alias, type = "MailContact"};
            bool result = CreateMailContact(ref contact);
            if (!result)
                throw new Exception(contact.error);
            else
                return result;
        }

        // CreateMailContact()
        // desc: Method creates a new mail contact, returns mail contact alias on success
        // params: string name  - Name of contact
        //         string email - External email of contact
        //         string ou    - Organizational Unit in which to create contact in
        // method: public
        // return: bool
        public bool CreateMailContact(ref ExchangeUser newContact)
        {
            if (newContact.alias == "") // If alias hasn't been set, give a default alias of the name minus any spaces
                newContact.alias = newContact.cn.Replace(" ", "");

            bool Result = false;
            string ErrorSet = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;
            ExchangeUser contact = GetMailContact(newContact);

            if (contact == null)
            {
                using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
                {
                    try
                    {
                        thisRunspace.Open();

                        using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                        {
                            try
                            {
                                thisPipeline.Commands.Add("New-MailContact");
                                thisPipeline.Commands[0].Parameters.Add("Name", @newContact.cn);
                                thisPipeline.Commands[0].Parameters.Add("ExternalEmailAddress", @newContact.email);
                                thisPipeline.Commands[0].Parameters.Add("OrganizationalUnit", @newContact.ou);
                                thisPipeline.Commands[0].Parameters.Add("Alias", @newContact.alias);
                                try
                                {
                                    thisPipeline.Invoke();
                                    Result = true;
                                }
                                catch (Exception ex)
                                {
                                    ErrorSet = "Error: " + ex.ToString();
                                    newContact.error += "Error: " + ex.ToString();
                                }
                                // Check for errors in the pipeline and throw an exception if necessary.
                                if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                                {
                                    StringBuilder pipelineError = new StringBuilder();
                                    pipelineError.AppendFormat("Error calling New-MailContact.");
                                    foreach (object item in thisPipeline.Error.ReadToEnd())
                                    {
                                        pipelineError.AppendFormat("{0}\n", item.ToString());
                                    }
                                    ErrorSet = ErrorSet + "Error: " + pipelineError.ToString();
                                    newContact.error += "Error: " + pipelineError.ToString();
                                }

                            }
                            catch (Exception ex)
                            {
                                ErrorSet = "Error: " + ex.ToString();
                                newContact.error += "Error: " + ex.ToString();
                            }
                        }
                    }
                    finally
                    {
                        thisRunspace.Close();
                    }
                }
            }
        

            if (contact != null && contact.email != newContact.email)
            {
                throw new Exception("Contact Conflict");
            }

            if (contact != null) // If at the end of everything
            {
                newContact.alias = contact.alias;
            }

            return Result;
        }

        /// <summary>
        /// Deletes mail contacts, useful in testing.
        /// </summary>
        /// <param name="alias"></param>
        public void DeleteMailContact(string alias)
        {
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;

            ExchangeUser contact = GetMailContact(alias);
            if (contact == null)
                return;

            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        thisPipeline.Commands.Add("Remove-MailContact");
                        thisPipeline.Commands[0].Parameters.Add("Identity", alias);
                        thisPipeline.Commands[0].Parameters.Add("Confirm", false);
                        thisPipeline.Invoke();

                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0)
                        {
                            StringBuilder pipelineError = new StringBuilder();
                            foreach (object item in thisPipeline.Error.ReadToEnd())
                            {
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            throw new Exception(pipelineError.ToString());
                        }
                    }
                }
                finally
                {
                    thisRunspace.Close();
                }
            }
        }

        #endregion

        public string GetIdentity()
        {
            AppDomain.CurrentDomain.SetPrincipalPolicy(System.Security.Principal.PrincipalPolicy.WindowsPrincipal);
            System.Security.Principal.WindowsPrincipal user = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
            return user.Identity.Name;
        }
    }
}