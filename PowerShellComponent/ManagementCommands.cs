using System;
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

// Scope PowerShellComponent
namespace PowerShellComponent
{
    // Class ManagementCommands
    public class ManagementCommands : System.EnterpriseServices.ServicedComponent
    {
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
                        thisPipeline.Commands[0].Parameters.Add("Database", @"mail2007.thetoolbox.com\First Storage Group\Mailbox Database");
                        thisPipeline.Commands[0].Parameters.Add("DomainController", "adtest2008.thetoolbox.com");
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
                            ReturnSet = GetUser(attributes["alias"].Replace("-vpn", ""));
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
                        thisPipeline.Commands[0].Parameters.Add("Database", @"mail2007.thetoolbox.com\First Storage Group\Mailbox Database");
                        thisPipeline.Invoke();
                        try{
                            ReturnSet = GetUser(attributes["alias"]);
                        }catch (Exception ex){
                            ErrorText = "Error: " + ex.ToString();
                            return ErrorText;
                        }
                        // Check for errors in the pipeline and throw an exception if necessary.
                        if (thisPipeline.Error != null && thisPipeline.Error.Count > 0){
                            StringBuilder pipelineError = new StringBuilder();
                            pipelineError.AppendFormat("Error calling New-MailUser.");
                            foreach (object item in thisPipeline.Error.ReadToEnd()){
                                pipelineError.AppendFormat("{0}\n", item.ToString());
                            }
                            ErrorText = ErrorText + "Error: " + pipelineError.ToString();
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
                        thisPipeline.Commands[0].Parameters.Add("DomainController", "adtest2008.thetoolbox.com");
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

        // GetUser()
        // desc: Method uses PowerShellSnapIn "Microsoft.Exchange.Management.PowerShell.Admin" to get user mailbox using CMDLET Get-Mailbox
        // params: string identity  - User login name
        //         int current_page - Current Page to return
        //         int per_page     - Entries to return per page
        // method: public
        // return: string, ExchangeUser XML serialized
        public string GetUser(string identity, int current_page = 1, int per_page = 10)
        {
            List<ExchangeUser> users = new List<ExchangeUser>();
            ExchangeUserShorter shorty = new ExchangeUserShorter() { CurrentPage = current_page, PerPage = per_page };
            int total_entries;

            try
            {
                users = GetUsers(out total_entries, identity:identity, displayName:"", current_page:current_page, per_page:per_page);
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

        private List<ExchangeUser> GetUsers(out int total_entries, string identity = "", string displayName = "", int current_page = 0, int per_page = 0)
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
                        if (identity.IndexOf("-vpn") != -1) thisPipeline.Commands.Add("Get-User");
                        else thisPipeline.Commands.Add("Get-Mailbox");
                        if (identity != "") thisPipeline.Commands[0].Parameters.Add("Identity", @identity);
                        if (displayName != "") thisPipeline.Commands[0].Parameters.Add("Anr", @displayName);
                        thisPipeline.Commands[0].Parameters.Add("SortBy", "DisplayName");
                        try
                        {
                            Collection<PSObject> original_results = thisPipeline.Invoke();
                            total_entries = original_results.Count;
                            IEnumerable<PSObject> results = null;

                            if (current_page == 0 && per_page == 0)
                                results = original_results;
                            else if (current_page < 2)
                                results = original_results.Take<PSObject>(per_page + 1);
                            else
                                results = original_results.Skip<PSObject>((current_page - 1) * per_page).Take<PSObject>(per_page);

                            foreach (PSObject result in results)
                            {
                                user = ReadUserInformation(result);

                                if (user.upn != "")
                                {
                                    using (Pipeline newPipeline = thisRunspace.CreatePipeline())
                                    {
                                        string vpn_identity = user.upn.Replace("@", "-vpn@");
                                        if (identity.IndexOf("-vpn") != -1) vpn_identity = user.upn;
                                        newPipeline.Commands.Add("Get-User");
                                        newPipeline.Commands[0].Parameters.Add("Identity", @vpn_identity);
                                        foreach (PSObject result2 in newPipeline.Invoke())
                                        {
                                            user.has_vpn = (((string)result2.Members["UserPrincipalName"].Value).Length > 0);
                                        }
                                    }
                                    users.Add(user);
                                }
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

        // CreateDistributionGroup()
        // desc: Method creates a new distribution list
        // params: string group_name - Name of new disitribution list
        // method: public
        // return: bool
        public string CreateDistributionGroup(string group_name, string ou)
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
                                DistributionGroupsShorter shorty = XmlSerializationHelper.Deserialize<DistributionGroupsShorter>(GetDistributionGroup(group_name, 0, 0));
                                if (shorty.groups.Count > 0)
                                    group = shorty.groups[0];
                                else
                                    throw new Exception("Group creation failed somewhere, new group was not found");
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
        public string GetDistributionGroup(string identity, int current_page, int per_page)
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
                            Collection<PSObject> original_results = thisPipeline.Invoke();
                            total_entries = original_results.Count;
                            IEnumerable<PSObject> results = null;
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
                                     
                                   }
                                     
                               }
                               if (group.displayName.Length > 0)
                               {
                                   group.users = new ExchangeUserMembers();
                                   //if(identity != "") 
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
                        user.ou = user.ou.Substring(user.ou.IndexOf('/') + 1);
                        break;
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
                return false;
        }

        public string UpdateDistributionGroup(string distributionGroupXml)
        {
            // We're going to update our users first. Othere things may follow, but users are the only one for now
            DistributionGroup group = XmlSerializationHelper.Deserialize<DistributionGroup>(distributionGroupXml);

            // get the current group member list
            List<ExchangeUser> currentMembers = GetDistributionGroupMembers(group.Name);
            
            // find the members that are new: those that don't exist in the currentMember list
            List<ExchangeUser> newUsers = group.users.users.Except(currentMembers, new LambdaComparer<ExchangeUser>((x, y) => x.alias == y.alias)).ToList();

            // find the members that are removed
            List<ExchangeUser> removedUsers = currentMembers.Except(group.users.users, new LambdaComparer<ExchangeUser>((x, y) => x.alias == y.alias)).ToList();

            // Add users using the existing methodology
            newUsers.FindAll(x => x.type == "MailContact").ForEach(x => CreateMailContact(ref x));
            newUsers.ForEach(x => AddToDistributionGroup(group.Name, x.alias));

            // Do the remove here, if we're removing anything
            removedUsers.ForEach(x => RemoveFromDistributionGroup(group.Name, x.alias));

            group.users.users = newUsers;
            
            // return the new group, though we won't use it.
            return XmlSerializationHelper.Serialize(XmlSerializationHelper.Deserialize<DistributionGroupsShorter>(GetDistributionGroup(group.Name, 0, 0)).groups[0]);
        }


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
        public bool CreateMailContact(ref ExchangeUser newContact, int limiter = 1)
        {
            if (newContact.alias == "") // If alias hasn't been set, give a default alias of the name minus any spaces
                newContact.alias = newContact.cn.Replace(" ", "");

            bool Result = false;
            string ErrorSet = "";
            RunspaceConfiguration config = RunspaceConfiguration.Create();
            PSSnapInException warning;
            config.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out warning);
            if (warning != null) throw warning;
            ExchangeUser contact = null;
            using (Runspace thisRunspace = RunspaceFactory.CreateRunspace(config))
            {
                try
                {
                    thisRunspace.Open();
                    
                    // first look for the contact to see if it already exists
                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        if (newContact.alias != "")
                        {
                            thisPipeline.Commands.Add("Get-Contact");
                            thisPipeline.Commands[0].Parameters.Add("Identity", @newContact.alias);
                            try
                            {
                                Collection<PSObject> results = thisPipeline.Invoke();

                                if (results.Count > 0)
                                    contact = ReadUserInformation(results[0]);
                            }
                            catch { } // We don't really care about what went wrong, we're just going to say that the contact does not exist
                        }
                    }

                    using (Pipeline thisPipeline = thisRunspace.CreatePipeline())
                    {
                        try
                        { 
                            if (contact == null)
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
                            else
                                Result = true;
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

            if (contact != null && contact.email != newContact.email)
            {
                // We found a user with the alias, but they had a different email, now we have to go unique alias hunting
                int aliasNumber = 1;
                if(int.TryParse(newContact.alias.Replace(newContact.cn.Replace(" ", ""), ""), out aliasNumber)) // replace, in the provided alias, the "default" alias to try to get a number
                {
                    aliasNumber++;
                }

                if (aliasNumber == 1 && limiter > 2 || limiter > 99) // Break out of the loop if the alias number isn't incrementing or if we've tried 99 other aliases
                {
                    return false;
                }
                else
                {
                    newContact.alias = newContact.cn.Replace(" ", "") + aliasNumber.ToString();
                    return this.CreateMailContact(ref newContact, limiter + 1); // recurse this method to try to find another contact.
                }
            }

            return Result;
        }

        public string GetIdentity()
        {
            AppDomain.CurrentDomain.SetPrincipalPolicy(System.Security.Principal.PrincipalPolicy.WindowsPrincipal);
            System.Security.Principal.WindowsPrincipal user = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
            return user.Identity.Name;
        }
    }
}