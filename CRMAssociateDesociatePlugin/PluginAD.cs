using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace CRMAssociateDesociatePlugin
{
    public class PluginAD:IPlugin
    {
        //
        public IPluginExecutionContext Context { get; private set; }
        public IOrganizationService Service { get; private set; }
        //
        public void CreateEmail(EntityReference contactRef, EntityReferenceCollection accountsReference,Guid currentUserId,string subject,string description,Guid regardingObjId)
        {
            Entity contact= Service.Retrieve(contactRef.LogicalName, contactRef.Id, new ColumnSet(new string[] { "contactid","fullname" }));
            var SendEmailAction = new OrganizationRequest()
            {
                RequestName = "new_SendCustomEmailAction"
            };
            SendEmailAction["Sender"] = new EntityReference("systemuser", currentUserId);
            SendEmailAction["RecepientEmail"] = new EntityReference("contact", (Guid)contact.Attributes["contactid"]);
            //
            foreach (var accountRef in accountsReference)
            {
                Entity account = Service.Retrieve(accountRef.LogicalName, accountRef.Id, new ColumnSet(new string[] {"name"}));
                var linkContact = $"https://andriikyrstiuksenvironment.crm11.dynamics.com/main.aspx?appid=42251675-59f8-ea11-a815-000d3a86b9ef&pagetype=entityrecord&etn=account&id=" + $"{account.Id}";
                SendEmailAction["Subject"] = $"{subject} {account.Attributes["name"]} with contact {contact.Attributes["fullname"]}";
                SendEmailAction["Body"] = $"{description} - {linkContact}";
                if(regardingObjId!=Guid.Empty)
                    SendEmailAction["RegardingContact"] = new EntityReference("contact", regardingObjId);
                Service.Execute(SendEmailAction);
            }
        }
        //
        public void Execute(IServiceProvider serviceProvider)
        {
            Context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            Service = factory.CreateOrganizationService(Context.UserId);
            //
            EntityReference targetEntity = null;
            EntityReferenceCollection relatedEntities = null;
            //
            var currentUser = Service.Retrieve("systemuser", Context.UserId, new ColumnSet("fullname"));
                //
                if (Context.MessageName == "Associate" || Context.MessageName == "Disassociate")
                {
                    if (Context.InputParameters.Contains("Relationship"))
                    {
                        var relationshipName = ((Relationship)Context.InputParameters["Relationship"]).SchemaName; ;
                        if (relationshipName != "cr53f_Contact_Account_Account") return;
                        //
                        if (Context.InputParameters.Contains("Target") && Context.InputParameters["Target"] is EntityReference)
                        {

                            targetEntity = (EntityReference)Context.InputParameters["Target"];
                        }
                        //
                        if (Context.InputParameters.Contains("RelatedEntities") && Context.InputParameters["RelatedEntities"] is EntityReferenceCollection)
                        {

                            relatedEntities = Context.InputParameters["RelatedEntities"] as EntityReferenceCollection;
                        }
                    }
                    tracingService.Trace("Event-Delegate Associate & Dissassociate Part 1 Successfully");
                    //
                    switch (Context.MessageName)
                    {
                        case "Associate":
                            CreateEmail(targetEntity, relatedEntities, currentUser.Id, "Account has been associated", "Account has been added", targetEntity.Id);
                            break;
                        case "Disassociate":
                            CreateEmail(targetEntity, relatedEntities, currentUser.Id, "Account has been disassociated", "Account has been removed", Guid.Empty);
                            break;
                    }
                }
        }
    }
}
