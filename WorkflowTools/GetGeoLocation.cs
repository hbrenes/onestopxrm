using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System.Collections;
using System.Dynamic;
using System.Web.Script.Serialization;
using System.Net;

namespace OneStopXRM.WorkflowTools
{
    public class GetGeoLocationByAddress : CodeActivity
    {
        #region  Parameters

        [Input("Street")]
        public InArgument<string> Street { get; set; }

        [Input("City")]
        public InArgument<string> City { get; set; }

        [Input("State")]
        public InArgument<string> State { get; set; }

        [Input("Zip Code")]
        public InArgument<string> ZipCode { get; set; }

        [Input("Google API Key")]
        public InArgument<string> APIKey { get; set; }

        [Output("Latitude")]
        public OutArgument<decimal> Latitude { get; set; }

        [Output("Longitude")]
        public OutArgument<decimal> Longitude { get; set; }


        #endregion

        protected override void Execute(CodeActivityContext context)
        {
            // Get the tracing service
            ITracingService tracingService =
            context.GetExtension<ITracingService>();

            //Create the context
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();

            // Get the context service.
            IWorkflowContext mycontext =
            context.GetExtension<IWorkflowContext>();

            IOrganizationServiceFactory serviceFactory =
            context.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of CrmService.
            IOrganizationService crmService =
            serviceFactory.CreateOrganizationService(mycontext.UserId);

            //instantiate input variables.
            string street = Street.Get(context);
            string city = City.Get(context);
            string state = State.Get(context);
            string zipCode = ZipCode.Get(context);
            string apiKey = APIKey.Get(context);

            string completeAddress = String.Format("{0}, {1}, {2} {3}", street, city, state, zipCode);

            //Geo code

            WebClient client = new WebClient();
            string value = client.DownloadString("https://maps.googleapis.com/maps/api/geocode/json?address=" + completeAddress + "&key=" + apiKey);
            
            JavaScriptSerializer jss = new JavaScriptSerializer();
            dynamic obj = jss.Deserialize<dynamic>(value);

            if (obj["status"] == "OK")
            {
                //set output latitude
                Latitude.Set(context, obj["results"][0]["geometry"]["location"]["lat"]);

                //set longitude
                Longitude.Set(context, obj["results"][0]["geometry"]["location"]["lng"]);
            }   
        }
    }
}