using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Web.Services.Protocols;

namespace MCS.ArtifactManagement
{
    /// <summary>    /// 
    /// Workflow  ProcessingOnChange  
    /// Registered OnUpdate for  Account, Incident, gnext_expense  for ITVERP
    /// 
    /// Workflow Input String specifiers used for dynamic use of additional entities that require Artifacts.
    /// 
    /// string  CaseLookupSchemaName =  the targetEntity Case lookup fieldname 
    /// string  ArtifactLookupSchemaName = the Artifact lookup fieldname for the targetEntity
    /// 
    /// Called to create Artifact records associated with target entity.
    /// Update existing Artifact records
    /// Delete nonApplicable Artifact records
    /// 
    /// Artifact Rule configuration records drive the Artifact record creation
    /// 
    /// </summary>
    public class ProcessingOnChange : CodeActivity
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
                tracer.Trace("ProcessingOnChange:Execute - Manage Creation/Update/Deletion of related Artifacts");

                // an exectute multiple of artifacts to update that runs after processing
                var artifactsToUpdate = new ExecuteMultipleRequest()
                {
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = true,
                        ReturnResponses = false
                    },
                    Requests = new OrganizationRequestCollection()
                };

                Entity entityTarget = (Entity)context.InputParameters["Target"];
                var fullEntity = service.Retrieve(entityTarget.LogicalName, entityTarget.Id, new ColumnSet(true));

                var incidentId = Guid.Empty;
                var contactId = Guid.Empty;
                var accountId = Guid.Empty;
                string artifactLookupName = "";

                switch (fullEntity.LogicalName)
                {
                    // Case is the Main entity for Artifact
                    case Incident.EntityLogicalName:
                        artifactLookupName = "mcs_caseid";
                        incidentId = fullEntity.Id;
                        contactId = fullEntity.Contains("primarycontactid") ? ((EntityReference)fullEntity["primarycontactid"]).Id : Guid.Empty;
                        accountId = fullEntity.Contains("customerid") ? ((EntityReference)fullEntity["customerid"]).Id : Guid.Empty;
                        break;

                    // Related Parent 
                    case Account.EntityLogicalName:
                        artifactLookupName = "mcs_accountid";
                        accountId = fullEntity.Id;
                        contactId = fullEntity.Contains("primarycontactid") ? ((EntityReference)fullEntity["primarycontactid"]).Id : Guid.Empty;
                       
                        using (var xrm = new CrmServiceContext(service))
                        {
                            // Find first created Child Incident to relate Artifact 
                            var a = xrm.IncidentSet.Where(x => x.CustomerId.Id == fullEntity.Id)
                                 .OrderBy(x => x.CreatedOn)
                                 .Select(x => new Incident()
                                 {
                                     Id = x.Id
                                 }).FirstOrDefault();

                            if (a != null)
                                incidentId = a.Id;
                        }
                        break;

                    // Related Parent
                    case Contact.EntityLogicalName:
                        artifactLookupName = "mcs_contactid";
                        contactId = fullEntity.Id;
                        accountId = fullEntity.Contains("parentcustomerid") ? ((EntityReference)fullEntity["parentcustomerid"]).Id : Guid.Empty;

                        using (var xrm = new CrmServiceContext(service))
                        {
                            // Find first created Child Incident to relate Artifact 
                            var c = xrm.IncidentSet.Where(x => x.PrimaryContactId.Id == entityTarget.Id)
                             .OrderBy(x => x.CreatedOn)
                             .Select(x => new Incident()
                             {
                                 Id = x.Id
                             }).FirstOrDefault();

                            if (c != null)
                                incidentId = c.Id;
                        }
                        break;

                    // Related Child
                    default:
                        // Retrieve WF arguments
                        artifactLookupName = ArtifactLookupSchemaName.Get(executionContext);
                        var caseLookup = CaseLookupSchemaName.Get(executionContext);

                        if (caseLookup == null || artifactLookupName == null) return;

                        tracer.Trace("ProcessingOnChange:Execute - Find parent Case using lookup " + caseLookup);
                        // if we have the proper workflow arguments retrieve and set the incidentid, accountid, and contactid associated
                        // with this related record
                        incidentId = ((EntityReference)fullEntity[caseLookup]).Id;                    
                                                
                        using (var xrm = new CrmServiceContext(service))
                        {
                            var e = (from c in xrm.IncidentSet where c.Id == incidentId
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

                // Find existing Artifacts related to the triggering Entity specified in the Artifact by the artifactLookupName
                tracer.Trace("ProcessingOnChange:Execute - Find Related Artifacts for "+ artifactLookupName);

                List<mcs_artifact> existingArtifacts;
                using (var xrm = new CrmServiceContext(service))
                {
                    existingArtifacts = xrm.mcs_artifactSet.Where(x => ((EntityReference)x[artifactLookupName]).Id == fullEntity.Id)
                       .Select(x => new mcs_artifact()
                       {
                           mcs_artifactid = x.mcs_artifactid,
                           mcs_ArtifactType = x.mcs_ArtifactType,
                           mcs_ArtifactRuleId = x.mcs_ArtifactRuleId,
                           mcs_Association = x.mcs_Association,
                           mcs_AccountId = x.mcs_AccountId,
                           mcs_Contactid = x.mcs_Contactid,
                           mcs_CaseId = x.mcs_CaseId
                       }).ToList();
                }

                tracer.Trace("ProcessingOnChange:Execute - initialize Artifact Service");
                var artifactService = new ArtifactService(service, fullEntity, artifactLookupName, accountId, contactId, incidentId);

                tracer.Trace("ProcessingOnCreation:CreateRequiredDocs Number of Artifact Rules =" + artifactService.ArtifactRules.Count());

                foreach(mcs_artifactrule r in artifactService.ArtifactRules)
                {
                    tracer.Trace("ProcessingOnCreation:CreateRequiredDocs Number of Found Rules Specifier =" + r.mcs_QuestionIdentifier + " = " + r.mcs_SuccessIndicator);
                }
                               
                var nonApplicableArtifactRuleIds = new List<Guid>();

                // Note:  entityTarget ONLY includes the fields that have just been modified
                // For any reconciliation we would need the use the fullEntity with all Attributes
                // and we would not want to filter out null Question Identifiers for anything missed during an OnCreate.
                //
                var processingItems = from a in artifactService.ArtifactRules.Where(x => x.mcs_QuestionIdentifier != null)
                                      join b in entityTarget.Attributes on a.mcs_QuestionIdentifier equals b.Key.ToString()
                                      select new { a, b };

                tracer.Trace("ProcessingOnCreation:CreateRequiredDocs Number of ProcessingItems =" + processingItems.Count());

                foreach (var item in processingItems)
                {
                    // if we don't have a question identifer move on to the next processing item because artifact rules with no 
                    // question identifier only matter in the context of a create event not update
                    if (item.b.Value == null) continue;                   

                    var attribute = item.b;
                    var artifactRule = item.a;
                    string attributeText = "";
                    string attributeName = "";
                    int attributeValue = new int();

                    attributeName = attribute.Key.ToString();
                    // we do not have artifact rules based on state or status reason, move to the next processing item
                    if (attributeName == "statuscode") continue;
                    if (attributeName == "statecode") continue;

                    // get the attribute text and value based on what type the attribute value is
                    switch (attribute.Value)
                    {
                        case OptionSetValue o:  // get the text value of selected item
                            attributeValue = o.Value;
                            attributeText = ArtifactService.GetOptionSetText(entityTarget.LogicalName, attributeName, attributeValue, service);
                            break;
                        case EntityReference eRef:
                        case null: continue;
                        default:
                            attributeText = attribute.Value.ToString();
                            break;
                    }

                    // Find any existing Artifacts against this rule.
                    IEnumerable<mcs_artifact> existingArtifactsforRule = existingArtifacts.Where(x => x.mcs_ArtifactRuleId.Id == artifactRule.Id);
                    var alreadyExists = existingArtifactsforRule.Count() > 0;

                    // C# 7.0 pattern matching.. This allows us to use both string and integer values of option sets
                    switch (artifactRule.mcs_SuccessIndicator)
                    {
                        // if we can parse from the string the success indicator is an optionsetvalue
                        case string s when int.TryParse(s, out int i):
                            // if the question identifier equals the success indicator and does not already exist create the artifact
                            if (attributeValue == i)
                            {
                                if (!alreadyExists)
                                {
                                    tracer.Trace("ProcessingOnChange:Execute - CreateArtifact text =" + attributeValue + "Specifier =" + i);
                                    artifactService.CreateArtifact(artifactRule);
                                    continue;
                                }
                            }
                            // if the question identifier does not equal the success indicator add to list of potential artifacts that need
                            // to be deleted
                            else
                            {
                                tracer.Trace("ProcessingOnChange:Execute - AddtoDeleteList text =" + attributeValue + "Specifier =" + i);
                                nonApplicableArtifactRuleIds.Add(artifactRule.Id);
                                continue;
                            }
                            break;
                        // if we have a string value in the success indicator
                        case string s:
                            // if the question identifier equals the success indicator and does not already exist create the artifact
                            if (attributeText == s)
                            {
                                if (!alreadyExists)
                                {
                                    tracer.Trace("ProcessingOnChange:Execute - CreateArtifact text =" + attributeText + "  Specifier =" + s);
                                    artifactService.CreateArtifact(artifactRule);
                                    continue;
                                }
                            }
                            // if the question identifier does not equal the success indicator add to list of potential artifacts that need
                            // to be deleted
                            else
                            {
                                tracer.Trace("ProcessingOnChange:Execute - AddtoDeleteList text =" + attributeText + "  Specifier =" + s);
                                nonApplicableArtifactRuleIds.Add(artifactRule.Id);
                                continue;
                            }

                            break;

                        default: // Skip to next Rule if this doesn't match
                            continue;
                    }

                    // If we didn't just create a new Artifact or add this Artifact Rule to the nonApplicable list, consider updating the Artifact if something has changed.

                    // Check if attribute is a specifier

                    if (attributeName == artifactRule.mcs_Specifier)
                    {
                        // loop through the collection of Artifacts for this rule
                        foreach (var artifact in existingArtifactsforRule)
                        {
                            // instantiate an update object for document in question
                            var updateArtifact = new mcs_artifact()
                            {
                                mcs_artifactid = artifact.Id
                            };

                            tracer.Trace("ProcessingOnChange:Execute - GetAssociationByRule ");
                            var association = artifactService.GetArtifactAssociationByRule(artifactRule);

                            tracer.Trace("ProcessingOnChange:Execute - Done ");

                            var dirty = false;

                            if (artifact.mcs_Association != association)
                            {
                                updateArtifact.mcs_Association = association;
                                dirty = true;
                            }

                            // Update Links if they don't already exist
                            if ((artifact.mcs_AccountId == null || artifact.mcs_AccountId.Id != accountId) && accountId != Guid.Empty)
                            {
                                updateArtifact.mcs_AccountId = new EntityReference("account", accountId);
                                dirty = true;
                            }
                            if ((artifact.mcs_Contactid == null || artifact.mcs_Contactid.Id != contactId) && contactId != Guid.Empty)
                            {
                                updateArtifact.mcs_Contactid = new EntityReference("contact", contactId);
                                dirty = true;
                            }                                                      

                            if ((artifact.mcs_CaseId == null || artifact.mcs_CaseId.Id != incidentId) && incidentId != Guid.Empty)
                            {
                                updateArtifact.mcs_CaseId = new EntityReference("incident", incidentId);
                                dirty = true;
                            }

                            if (dirty)
                            {
                                var updateReq = new UpdateRequest()
                                {
                                    Target = updateArtifact
                                };

                                artifactsToUpdate.Requests.Add(updateReq);
                            }
                        }
                    }

                    // Update any changed Artifacts
                    service.Execute(artifactsToUpdate);
                }
                tracer.Trace("Artifact ProcessingOnChange:Execute - Delete any related Artifacts that are no longer applicable");
                         
                // for matches between the existing artifacts and our nonapplicable artifact rules delete the artifact
                (from a in existingArtifacts
                 join b in nonApplicableArtifactRuleIds on a.mcs_ArtifactRuleId.Id equals b
                 select a.Id).ToList().ForEach(x => service.Delete(mcs_artifact.EntityLogicalName, x));
                
               // throw (new Exception("testing"));
            }
            catch (SoapException e)
            {
                throw new InvalidPluginExecutionException(e.Detail.InnerXml);
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
    }
}