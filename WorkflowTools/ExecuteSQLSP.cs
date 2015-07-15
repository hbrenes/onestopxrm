using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.ApplicationBlocks.Data;
using System.Data;

namespace OneStopXRM.WorkflowTools
{
    public class ExecuteSQLSP : CodeActivity
    {
        [Input("SP Name")]
        public InArgument<string> inSPName { get; set; }
        [Input("SP Params")]
        public InArgument<string> inSPParams { get; set; }
        [Input("SP DBInstance")]
        public InArgument<string> inSPDBInstance { get; set; }
        [Input("SP Database")]
        public InArgument<string> inSPDatabase { get; set; }
        [Input("Username")]
        public InArgument<string> inUsername { get; set; }
        [Input("Password")]
        public InArgument<string> inPassword { get; set; }

        [Output("SPResult")]
        public OutArgument<string> SPResult { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            //initialization code
            try
            {
                // Get the tracing service
                ITracingService tracingService =
                context.GetExtension<ITracingService>();

                // Get the context service.
                IWorkflowContext mycontext =
                context.GetExtension<IWorkflowContext>();

                IOrganizationServiceFactory serviceFactory =
                context.GetExtension<IOrganizationServiceFactory>();

                // Use the context service to create an instance of CrmService.
                IOrganizationService crmService =
                serviceFactory.CreateOrganizationService(mycontext.UserId);

                //Split parameters by Pipe
                string[] parameters = inSPParams.Get(context).Split('|');

                //create object parameters for sqlhelper
                object[] sqlPars = new object[parameters.Length];

                for (int i = 0; i < parameters.Length; i++)
                {
                    sqlPars[i] = parameters[i];
                }

                //format connection string
                string conn = String.Format("Data Source={0};Initial Catalog={1};User Id={2};Password={3}",
                                            inSPDBInstance.Get(context),
                                            inSPDatabase.Get(context),
                                            inUsername.Get(context),
                                            inPassword.Get(context));

                //Execute DataSet
                DataSet ds = SqlHelper.ExecuteDataset(conn, inSPName.Get(context), sqlPars);

                if (ds != null)
                {
                    if (ds.Tables.Count > 0)
                    {
                        SPResult.Set(context, ds.Tables[0].Rows[0][0].ToString());
                    }
                    else
                        SPResult.Set(context, "0");
                }
                else
                {
                    SPResult.Set(context, "0");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
