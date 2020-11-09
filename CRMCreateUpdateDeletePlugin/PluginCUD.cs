using Microsoft.Xrm.Sdk;
using System;

namespace CRMCreateUpdateDeletePlugin
{
    public class PluginCUD : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            ITracingService tracingService =(ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);
            var SendEmailAction = new OrganizationRequest()
            {
                RequestName = "new_SendCustomEmailAction"
            };
            //
            if (context.InputParameters.Contains("Target"))
            {
                Entity email = new Entity("email");
                Guid contactId;
                string LogicalName = string.Empty;
                object fullName= new object();
                object createdOn= new object();
                object emailAddress= new object();
                object modifiedOn = new object();
                if (context.InputParameters["Target"] is Entity)
                {
                    Entity contact = (Entity)context.InputParameters["Target"];
                    contactId = (Guid)contact.Attributes["contactid"];
                    LogicalName = contact.LogicalName;
                    if (!context.MessageName.Contains("Update"))
                    {
                        fullName = contact.Attributes["fullname"];
                        createdOn = contact.Attributes["createdon"];
                    }
                    else 
                    { 
                        emailAddress = contact.Attributes["emailaddress1"];
                        modifiedOn = contact.Attributes["modifiedon"];
                    }
                }
                else if (context.InputParameters["Target"] is EntityReference)
                {
                    EntityReference entityref = (EntityReference)context.InputParameters["Target"];
                    contactId = entityref.Id;
                    LogicalName = entityref.LogicalName;
                }
                tracingService.Trace("Project Id is {0}", contactId);
                SendEmailAction["Sender"] = new EntityReference("systemuser", context.UserId);
                SendEmailAction["RecepientEmail"] = new EntityReference("contact", contactId);
                if (LogicalName == "contact")
                {
                    switch (context.MessageName)
                    {
                        case "Create":
                            var linkContact = $"https://andriikyrstiuksenvironment.crm11.dynamics.com/main.aspx?appid=42251675-59f8-ea11-a815-000d3a86b9ef&pagetype=entityrecord&etn=contact&id=" + $"{contactId}";
                            SendEmailAction["Subject"] = $"New Contact {fullName} created {createdOn}";
                            SendEmailAction["Body"] = $"New contact created - {linkContact}";
                            SendEmailAction["RegardingContact"] = new EntityReference("contact", contactId);
                            break;
                        case "Update":
                            if (context.PreEntityImages.Contains("UpdatedEntity") && context.PreEntityImages["UpdatedEntity"] is Entity)
                            {
                                Entity preMessageImage = context.PreEntityImages["UpdatedEntity"];
                                SendEmailAction["Subject"] = $"Contact {preMessageImage.Attributes["fullname"]} email address changed {modifiedOn}";
                                SendEmailAction["Body"] = $"Old email address - {preMessageImage.Attributes["emailaddress1"]} <br> New email address {emailAddress}";
                                SendEmailAction["RegardingContact"] = new EntityReference("contact", contactId);
                            }
                            break;
                        case "Delete":
                            if (context.PreEntityImages.Contains("DeletedEntity") && context.PreEntityImages["DeletedEntity"] is Entity)
                            {
                                Entity preMessageImage = context.PreEntityImages["DeletedEntity"];
                                SendEmailAction["Subject"] = $"Contact {preMessageImage.Attributes["fullname"]} was deleted {preMessageImage.Attributes["modifiedon"]}";
                                SendEmailAction["Body"] = $"Contact was deleted!";
                            }
                            break;
                    }
                    service.Execute(SendEmailAction);
                }
            }
        }
    }
}
