using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using Microsoft.Xrm.Sdk.Query;
using System.Linq;


namespace MCS.ArtifactManagement
{
    /// <summary>
    /// Workflow  ProcessingOnCreation  
    /// Registered OnCreate for  Account, Incident, gnext_expense  for ITVERP
    /// 
    /// Workflow Input String specifiers used for dynamic use of additional entities that require Artifacts.
    /// 
    /// string  CaseLookupSchemaName =  the targetEntity Case lookup fieldname 
    /// string  ArtifactLookupSchemaName = the Artifact lookup fieldname for the targetEntity
    /// 
    /// Called to create Artifact records associated with target entity.
    /// Artifact Rule configuration records drive the Artifact record creation
    /// 
    /// </summary>
    public class ProcessingOnCreation : CodeActivity
    {
       
        [Input("Schema name for related record case lookup")]
        [Default(null)]
        public InArgument<string> CaseLookupSchemaName { get; set; }

        [Input("Schema name for related record artifact lookup")]
        [Default(null)]
        public InArgument<string> ArtifactLookupSchemaName { get; set; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                var targetName = context.PrimaryEntityName;
               
                tracer.Trace("ProcessingOnCreation:Execute retrieve triggering entity to process Artifacts for " + targetName);

                // retrieve all attributes from triggering entities because any field can be configured as a Question Identifier
                var targetEntity = service.Retrieve(targetName, context.PrimaryEntityId, new ColumnSet(true));
                
                var accountId = Guid.Empty;
                var contactId = Guid.Empty;
                var incidentId = Guid.Empty;
                var artifactLookupName = "";

                switch (targetName)
                {

                    // Assuming that the Parent Account and Contact will be created in advance of any incident.

                    case "account":
                        artifactLookupName = "mcs_accountid";
                        accountId = targetEntity.Id;
                        contactId = targetEntity.Contains("primarycontactid") ? ((EntityReference)targetEntity["primarycontactid"]).Id : Guid.Empty;
                        break;

                    case "contact":
                        artifactLookupName = "mcs_contactid";
                        accountId = targetEntity.Contains("parentcustomerid") ? ((EntityReference)targetEntity["parentcustomerid"]).Id : Guid.Empty;
                        var account = service.Retrieve("account", ((EntityReference)targetEntity["parentcustomerid"]).Id, new ColumnSet(new[] { "primarycontactid" }));
                        contactId = account.Contains("primarycontactid") ? ((EntityReference)account["primarycontactid"]).Id : Guid.Empty;
                        break;

                    case "incident":
                        incidentId = targetEntity.Id;
                        artifactLookupName = "mcs_caseid";
                        accountId = targetEntity.Contains("customerid") ? ((EntityReference)targetEntity["customerid"]).Id : Guid.Empty;
                        contactId = targetEntity.Contains("primarycontactid") ? ((EntityReference)targetEntity["primarycontactid"]).Id : Guid.Empty;
                        break;

                    // Below handles custom related records based on workflow input parameters
                    default:
                        // Retrieve WF arguments
                        artifactLookupName = ArtifactLookupSchemaName.Get(executionContext);
                        var caseLookup = CaseLookupSchemaName.Get(executionContext);

                        if (caseLookup == null || artifactLookupName == null) return;

                        // if we have the proper workflow arguments retrieve and set the incidentid, accountid, and contactid associated
                        // with this related record
                        incidentId = ((EntityReference)targetEntity[caseLookup]).Id;
                        using (var xrm = new CrmServiceContext(service))
                        {
                            var e = (from c in xrm.IncidentSet
                                     where c.Id == incidentId
                                     select new Incident()
                                     {
                                         PrimaryContactId = c.PrimaryContactId,
                                         CustomerId = c.CustomerId
                                     }).FirstOrDefault();

                            accountId = e.CustomerId != null ? e.CustomerId.Id : Guid.Empty;
                            contactId = e.PrimaryContactId != null ? e.PrimaryContactId.Id : Guid.Empty;
                        }
                        break;
                }

                // Instantiate helper object
                var artifactService = new ArtifactService(service, targetEntity, artifactLookupName, accountId, contactId, incidentId);
                              
                // Create Artifact based on target Entity and Artifact Rule 
                foreach (var rule in artifactService.ArtifactRules)
                {
                    // if the question identifier is null then in the context of a create this is a what is called a static rule
                    // which means that the artifact gets created regardless of any criteria on create of the targetEntity
                    if (rule.mcs_QuestionIdentifier == null)
                    {
                        tracer.Trace("ProcessingOnCreation:CreateRequiredDocs create new artifact with no Question Identifier");
                        artifactService.CreateArtifact(rule);
                    }
                    else if (targetEntity.Contains(rule.mcs_QuestionIdentifier))
                    {
                        tracer.Trace(targetEntity.LogicalName + "Check This Question Identifier " + rule.mcs_QuestionIdentifier + " against " + rule.mcs_SuccessIndicator);

                        string attr;
                        string attrLabel = "";
                        
                        // get the appropriate attribute values based on what type of field the question identifier is
                        if (targetEntity[rule.mcs_QuestionIdentifier] is OptionSetValue)
                        {
                            attr = ((OptionSetValue)targetEntity[rule.mcs_QuestionIdentifier]).Value.ToString();
                            attrLabel = ArtifactService.GetOptionSetText(targetEntity.LogicalName, rule.mcs_QuestionIdentifier, Convert.ToInt32(attr), service);
                        }
                        else if (targetEntity[rule.mcs_QuestionIdentifier] is bool)
                            attr = ((bool)targetEntity[rule.mcs_QuestionIdentifier]).ToString();
                        else
                            attr = (string)targetEntity[rule.mcs_QuestionIdentifier];

                        tracer.Trace("attr = " + attr);

                        // if the attribute value is equal to the success indicator add the required doc to the context for creation
                        if (attr == rule.mcs_SuccessIndicator || attrLabel == rule.mcs_SuccessIndicator)
                        {
                            artifactService.CreateArtifact(rule);
                        }
                    }
                } 
                 
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }       
    }
}