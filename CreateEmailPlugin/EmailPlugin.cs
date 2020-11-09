using Microsoft.Xrm.Sdk;
using System;

namespace CreateEmailPlugin
{
    public class EmailPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var service = factory.CreateOrganizationService(context.UserId);
            //throw new InvalidPluginExecutionException(context.InputParameters["Subject"].ToString());
            //throw new InvalidPluginExecutionException(context.MessageName);
            Entity toActivityParty = new Entity("activityparty");
            Entity fromActivityParty = new Entity("activityparty");
            fromActivityParty["partyid"] = context.InputParameters["Sender"];
            toActivityParty["partyid"] = context.InputParameters["RecepientEmail"];
            //
            Entity email = new Entity("email");
            email.Attributes["to"] = new Entity[] { toActivityParty };
            email.Attributes["from"] = new Entity[] { fromActivityParty };
            //
            email.Attributes["subject"] ="CE."+context.InputParameters["Subject"];
            email.Attributes["description"] = context.InputParameters["Body"];
            if (context.InputParameters["RegardingContact"] != null)
                email.Attributes["regardingobjectid"] = context.InputParameters["RegardingContact"];
            //CreatedEmail.Set(context, 
            service.Create(email);
        }
    }
}
