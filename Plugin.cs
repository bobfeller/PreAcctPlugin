using System;
using System.Net.Mail;
using System.Linq;
using System.Collections.Generic;
// Microsoft Dynamics CRM namespace(s)
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;


namespace ASG.Crm.Sdk.Web
{
    public class AcctPreCreateHandler : IPlugin
    {
        public IOrganizationService wService;
        /// <summary>
        /// This plugin does a precheck of contact domain name and determines whether account exists
        /// </summary>
        public void Execute(IServiceProvider serviceProvider)
        {
            //throw new InvalidPluginExecutionException("IN PRE");
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Verify we have an entity to work with
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {

                try
                {
                    // Obtain the target business entity from the input parmameters.
                    Entity entity = (Entity)context.InputParameters["Target"];
                    bool isUpdate = false;
                    // Verify that the entity represents an account.
                    if (entity.LogicalName == "account")
                    {
                        if (entity.Contains("accountid")) isUpdate = true;

                        // If originatingleadid is set, display
                        // Use "contains" because the indexer will throw if the column is not found
                        if (entity.Contains("originatingleadid") == true)
                        {
                            // Get orginating lead 
                            Guid lkLead = new Guid();
                            EntityReference erLead = ((EntityReference)entity["originatingleadid"]);
                            lkLead = erLead.Id;

                            // No need to continue if empty leadid guid
                            if (lkLead == Guid.Empty) return;

                            // Get Domain from email from lead
                            string strLeadDomain = GetEmailDomain(entity["emailaddress1"].ToString());

                            // if valid email/domain set UPNSuffix
                            if (!String.IsNullOrEmpty(strLeadDomain))
                            {
                                //String strUpnSuffix = new String("asg_upnsuffix", strLeadDomain);
                                entity.Attributes.Add("asg_upnsuffix", strLeadDomain);
                            }

                        }
                        // Obtain the organization service reference.
                        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                        wService = serviceFactory.CreateOrganizationService(context.UserId);

                        // Verify no dupes on account number
                        if (entity.Attributes.Contains("accountnumber"))
                        {
                            if (VerifyAcct(entity["accountnumber"].ToString(), wService, isUpdate))
                            {
                                //if (!entity.Properties.Contains("asg_addtogp"))
                                throw new InvalidPluginExecutionException("This account number already exists.");
                            }
                        }

                        // Check for Credit Limit/Hold change
                        if (isUpdate)
                        {
                            bool bUpdate = false;
                            if (CheckCredit(((Guid)entity["accountid"]), wService, true))
                            {
                                if (entity.Attributes.Contains("creditlimit"))
                                    if (mCreditLimit.Value != ((Money)entity["creditlimit"]).Value)
                                        bUpdate = true;
                                if (entity.Attributes.Contains("creditonhold"))
                                    if (bCreditHold != ((Boolean)entity["creditonhold"]))
                                        bUpdate = true;
                                if (bUpdate)
                                {
                                    DateTime now = DateTime.Now;
                                    //CrmDateTimeProperty dt = new CrmDateTimeProperty("asg_creditcheckdate", now);
                                    entity.Attributes.Add("asg_creditcheckdate", now);
                                }
                            }
                        }
                        else    // Insert operation (create)
                        {
                            if (entity.Attributes.Contains("creditlimit") || entity.Attributes.Contains("creditonhold"))
                            {
                                bool bUpdate = false;
                                if (entity.Contains("creditlimit"))
                                    if (((Money)entity["creditlimit"]).Value != 0.00m)
                                        bUpdate = true;
                                if (entity.Attributes.Contains("creditonhold"))
                                    if (((Boolean)entity["creditonhold"]) == true)
                                        bUpdate = true;
                                if (bUpdate)
                                {
                                    DateTime now = DateTime.Now;
                                    //CrmDateTimeProperty dt = new CrmDateTimeProperty("asg_creditcheckdate", now);
                                    entity.Attributes.Add("asg_creditcheckdate", now);
                                }
                            }
                        }
                        // Check for ownership change
                        string origCSREmail = "";
                        string newCSREmail = "";
                        string origName = "";
                        string newName = "";
                        if (isUpdate && entity.Contains("ownerid"))
                        {
                            // Get original owner
                            if (!GetOwnerInfo((Guid)entity["accountid"], wService, ref origCSREmail, ref origName, false, new Guid())) return;
                            // Get new owner
                            EntityReference refNewOwner = (EntityReference)entity["ownerid"];
                            if (!GetOwnerInfo((Guid)entity["accountid"], wService, ref newCSREmail, ref newName, true, refNewOwner.Id)) return;
                            // Get the Acct Name
                            string strAcctName = GetAcctName((Guid)entity["accountid"], wService);

                            // Determine if we neeed to query the db to get CustomerTypeCode
                            string CustomerTypeCodeText = "";
                            if (!entity.Contains("customertypecode"))
                                CustomerTypeCodeText = GetAccountToBeUpdated(entity.Id);
                            else
                            {
                                CustomerTypeCodeText = (entity.FormattedValues["customertypecode"]).ToString();
                            }
                            // Send an email out if ownership changed
                            if ((newCSREmail != origCSREmail) && (!string.IsNullOrEmpty(strAcctName)))
                            {
                                string strMsg = "Info: The account '" + strAcctName + "' has been re-assigned to " +
                                    newName + " from " + origName + ".\n\n" +
                                    "You may view this account at: <" + Properties.Settings.Default.AcctUrl + "{" + ((Guid)entity["accountid"]).ToString() + "}>";
                                SendMail(origCSREmail, newCSREmail, "CRM Account Ownership change (" + CustomerTypeCodeText + ")", strMsg);
                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("This account number already exists"))
                        throw;
                    else
                        return;
                }
            }
        }
        
        // Return the Customer Category text for existing quotes only (new quotes will carry the code through the entity reference)
        public string GetAccountToBeUpdated(Guid guidAccountId)
        {
            //get the Quote to be canceled in same Opportunity, exclude the wonQuote
            myServiceContext context = new myServiceContext(wService);  //This comes from GeneratedCodeWithContext.cs
            IQueryable<Account> dbAccount = from x in context.AccountSet
                                    where x.AccountId == guidAccountId
                                    select x;

            if (dbAccount.ToList().Count > 0)
                return ((Account)dbAccount).FormattedValues["CustomerTypeCode"].ToString();
            else
                return null;
        }

        public bool SendMail(string To, string CC, string Subject, string Message)
        {
            //return true;
            try
            {
                if (String.IsNullOrEmpty(To) || String.IsNullOrEmpty(CC)) return true;
                if (To.ToLower() == CC.ToLower()) return true;

                string strSubject = Subject;
                string strBody = Message;

                // Create the Mail Addresses and Mail Message object to send
                MailAddress maFrom = new MailAddress("do_not_reply@virtual.com", "CRM Notification");
                MailAddress maTo = new MailAddress(To);
                MailAddress maCC = new MailAddress(CC);
                //MailAddress maCC2 = new MailAddress("development@virtual.com");
                MailMessage message = new MailMessage(maFrom, maTo);
                message.CC.Add(maCC);
                message.Bcc.Add("development@virtual.com");
                message.Subject = strSubject;
                message.Body = strBody;
                message.IsBodyHtml = false;
                // Send the mail
                SmtpClient client = new SmtpClient("smtp.virtual.com");
                client.Send(message);
                return true;
            }
            catch (Exception ex)
            {
                // Display the error causing the email to be sent....
                throw new Exception("Unable to send email notification: " + ex.Message);
            }
        }

        // This method will parse the domain name from the email address
        private string GetEmailDomain(string emailaddress)
        {
            string[] strtemp = emailaddress.Split('@');
            if (strtemp.Length < 2)
                throw new InvalidPluginExecutionException("Invalid Email Address: " + emailaddress);
            string[] strtemp2 = strtemp[1].Split('.');
            if (strtemp2.Length < 2)
                throw new InvalidPluginExecutionException("Invalid Email Address: " + emailaddress);

            // Found domain - return
            return strtemp[1];
        }

        // This method will look up a Account based on the account id and return the name
        public static string GetAcctName(Guid accountid, IOrganizationService service)
        {
            try
            {
                QueryExpression query = new QueryExpression();

                // Set the query to retrieve User records.

                query.EntityName = "account";

                // Create a set of columns to return.
                // Create a ColumnSet and set attributes we want to retrieve
                ColumnSet cols = new ColumnSet(new string[] { "name" });

                // Create the ConditionExpressions.
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = "accountid";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(accountid);

                // Builds the filter based on the condition
                FilterExpression filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.And;
                filter.Conditions.Add(condition);

                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the values from Microsoft CRM.
                EntityCollection retrieved = service.RetrieveMultiple(query);

                if (retrieved.Entities.Count == 1)
                {
                    Account acct = (Account)retrieved.Entities[0];
                    return acct.Name.ToString();
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public static Boolean bCreditHold;
        public static Money mCreditLimit = new Money();
        // This method will look up a Account based on the account id
        public static bool CheckCredit(Guid accountid, IOrganizationService service, bool isUpdate)
        {
            try
            {
                QueryExpression query = new QueryExpression();

                // Set the query to retrieve User records.

                query.EntityName = "account";

                // Create a set of columns to return.
                // Create a ColumnSet and set attributes we want to retrieve
                ColumnSet cols = new ColumnSet(new string[] { "creditlimit", "creditonhold" });



                // Create the ConditionExpressions.
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = "accountid";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(accountid);

                // Builds the filter based on the condition
                FilterExpression filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.And;
                filter.Conditions.Add(condition);
                //filter.Conditions.Add(condition);
                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the values from Microsoft CRM.
                EntityCollection retrieved = service.RetrieveMultiple(query);

                if (retrieved.Entities.Count == 1)
                {
                    Account acct = (Account)retrieved.Entities[0];
                    bCreditHold = acct.CreditOnHold.Value;
                    try
                    {
                        mCreditLimit.Value = acct.CreditLimit.Value;
                    }
                    catch
                    {
                        mCreditLimit.Value = 0m;
                    }
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        // This method will look up a Account based on the account number
        public static bool VerifyAcct(string accountnumber, IOrganizationService service, bool isUpdate)
        {
            try
            {
                QueryExpression query = new QueryExpression();

                // Set the query to retrieve User records.

                query.EntityName = "account";

                // Create a set of columns to return.
                // Create a ColumnSet and set attributes we want to retrieve
                ColumnSet cols = new ColumnSet(new string[] { "accountnumber" });

                // Create the ConditionExpressions.
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = "accountnumber";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(accountnumber);

                // Builds the filter based on the condition
                FilterExpression filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.And;
                filter.Conditions.Add(condition);
                //filter.Conditions.Add(condition);
                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the values from Microsoft CRM.
                EntityCollection retrieved = service.RetrieveMultiple(query);

                if (retrieved.Entities.Count == 1 && isUpdate)
                {
                    return false;
                }
                if (retrieved.Entities.Count > 0)
                {
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        // Get original owner from acount
        public static bool GetOwnerInfo(Guid accountid, IOrganizationService service, ref string CSREmail, ref string OwnerName, bool NewOwner, Guid owner)
        {
            try
            {
                Guid assignedCSR = new Guid();

                CSREmail = string.Empty;
                OwnerName = string.Empty;

                Guid ownerGuid = new Guid();
                QueryExpression query = new QueryExpression();

                // Set the query to retrieve User records.
                query.EntityName = "account";

                // Create a set of columns to return.
                // Create a ColumnSet and set attributes we want to retrieve
                ColumnSet cols = new ColumnSet(new string[] { "ownerid" });


                // Create the ConditionExpressions.
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = "accountid";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(accountid);

                // Builds the filter based on the condition
                FilterExpression filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.And;
                filter.Conditions.Add(condition);
                //filter.Conditions.Add(condition);
                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the values from Microsoft CRM.
                if (!NewOwner)
                {
                    EntityCollection retrieved = service.RetrieveMultiple(query);

                    if (retrieved.Entities.Count == 1)
                    {
                        Account acct = (Account)retrieved.Entities[0];
                        try
                        {
                            ownerGuid = acct.OwnerId.Id;
                        }
                        catch
                        {
                            return false;
                        }

                    }
                }
                else
                {
                    ownerGuid = (Guid)owner;
                }
                // Get the owner
                query = new QueryExpression();

                // Set the query to retrieve Owner's User record.
                query.EntityName = "systemuser";

                // Create a set of columns to return.
                // Create a ColumnSet and set attributes we want to retrieve
                cols = new ColumnSet(new string[] { "asg_assignedcsrid", "fullname" });

                // Create the ConditionExpressions.
                condition = new ConditionExpression();
                condition.AttributeName = "systemuserid";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(ownerGuid);

                // Builds the filter based on the condition
                filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.And;
                filter.Conditions.Add(condition);
                //filter.Conditions.Add(condition);
                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the Orginal Owner Info
                EntityCollection users = service.RetrieveMultiple(query);
                if (users.Entities.Count > 0)
                {
                    if (users.Entities.Count > 0)
                    {
                        SystemUser user = (SystemUser)users.Entities[0];
                        if (user.Attributes.Contains("asg_assignedcsrid"))
                            assignedCSR = (Guid)((EntityReference)user["asg_assignedcsrid"]).Id;
                        else
                            return false;

                        if (assignedCSR == Guid.Empty) return false;
                        CSREmail = GetCSREmail(assignedCSR, service);
                        OwnerName = user["fullname"].ToString();
                    }
                }
                else
                {
                    return false;
                }

                return true;
                 
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            return true;
        }
        public static string GetCSREmail(Guid systemuserId, IOrganizationService wService)
        {
            try
            {
                QueryExpression query = new QueryExpression();

                // Set the query to retrieve Contact records.
                query.EntityName = "systemuser";

                // Create a set of columns to return.
                ColumnSet cols = new ColumnSet(new string[] { "internalemailaddress" });

                // Create the ConditionExpressions.
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = "systemuserid";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(systemuserId);

                // Builds the filter based on the condition
                FilterExpression filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.And;
                filter.Conditions.Add(condition);

                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the values from Microsoft CRM.
                EntityCollection retrieved = wService.RetrieveMultiple(query);

                if (retrieved.Entities.Count == 1)
                {
                    SystemUser conResult = (SystemUser)retrieved.Entities[0];
                    return conResult.InternalEMailAddress;
                }
                else
                {
                    throw new InvalidPluginExecutionException("User not found in CRM database.");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
      
    }
}
