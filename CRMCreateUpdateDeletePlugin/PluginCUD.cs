using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

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
            if (context.InputParameters.Contains("Target"))
            {
                WhoAmIRequest systemUserRequest = new WhoAmIRequest();
                WhoAmIResponse systemUserResponse = (WhoAmIResponse)service.Execute(systemUserRequest);
                Guid userId = systemUserResponse.UserId;
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
                //tracingService.Trace(fullName.ToString());
                //tracingService.Trace(modifiedOn.ToString());
                //tracingService.Trace(emailAddress.ToString());
                //
                //
                Entity toActivityParty = new Entity("activityparty");
                Entity fromActivityParty = new Entity("activityparty");
                //var contactId = (Guid)contact.Attributes["contactid"];
                fromActivityParty["partyid"] = new EntityReference("systemuser", userId);
                toActivityParty["partyid"] = new EntityReference("contact", contactId);
                    // Убедимся, что целевой объект представляет собой случай.
                    // Если нет, этот плагин был зарегистрирован неправильно
                if (LogicalName == "contact")
                {
                    email.Attributes["to"] = new Entity[] { toActivityParty };
                    email.Attributes["from"] = new Entity[] { fromActivityParty };
                    switch (context.MessageName)
                    {
                        case "Create":
                            var linkContact = $"https://andriikyrstiuksenvironment.crm11.dynamics.com/main.aspx?appid=42251675-59f8-ea11-a815-000d3a86b9ef&pagetype=entityrecord&etn=contact&id=" + $"{contactId}";
                            email.Attributes["subject"] = $"New Contact {fullName} created {createdOn}";
                            email.Attributes["description"] = $"New contact created - {linkContact}";
                            email.Attributes["regardingobjectid"] = new EntityReference("contact", contactId);
                            break;
                        case "Update":
                            if (context.PreEntityImages.Contains("UpdatedEntity") && context.PreEntityImages["UpdatedEntity"] is Entity)
                            {
                                Entity preMessageImage = context.PreEntityImages["UpdatedEntity"];
                                email.Attributes["subject"] = $"Contact {preMessageImage.Attributes["fullname"]} email address changed {modifiedOn}";
                                email.Attributes["description"] = $"Old email address - {preMessageImage.Attributes["emailaddress1"]} <br> New email address {emailAddress}";
                                email.Attributes["regardingobjectid"] = new EntityReference("contact", contactId);
                                //throw new InvalidPluginExecutionException("test1");
                            }
                           // else { throw new InvalidPluginExecutionException("test2"); }
                            break;
                        case "Delete":
                            if (context.PreEntityImages.Contains("DeletedEntity") && context.PreEntityImages["DeletedEntity"] is Entity)
                            {
                                Entity preMessageImage = context.PreEntityImages["DeletedEntity"];
                                email.Attributes["subject"] = $"Contact {preMessageImage.Attributes["fullname"]} was deleted {preMessageImage.Attributes["modifiedon"]}";
                                email.Attributes["description"] = $"Contact was deleted!";
                            }
                            break;
                    }
                    var emailId = service.Create(email);
                }
            }
        }
    }
}
