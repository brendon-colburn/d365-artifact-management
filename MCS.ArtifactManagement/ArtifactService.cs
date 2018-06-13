using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;


namespace MCS.ArtifactManagement
{
    /// <summary>
    /// Artifact Service is a helper class where we hold multiple functionalities that are used across workflow activities and plugins
    /// </summary>
    class ArtifactService
    {
        private IOrganizationService _orgService;
        private Entity _entityTarget;
        public List<mcs_artifactrule> ArtifactRules;       
        public string ArtifactLookupName;
        public Guid AccountId;
        public Guid ContactId;
        public Guid IncidentId;

        /// <summary>
        /// On instantiation set the properties of the Artifact Service for later use
        /// </summary>
        /// <param name="service">the org service</param>
        /// <param name="target">the target entity our logic is contextually running on</param>
        /// <param name="lookupName">the lookup schema name to the artifact entity</param>
        /// <param name="aId">the accountid associated with the case</param>
        /// <param name="cId">the contact id associated with the case</param>
        /// <param name="iId">the case/incident id</param>
        public ArtifactService(IOrganizationService service, Entity target, string lookupName, Guid aId, Guid cId, Guid iId)
        {
            _orgService = service;
            _entityTarget = target;
            ArtifactRules = (target.LogicalName == "incident") ?  GetArtifactRules():GetArtifactRuleRelatedRecord();          
            ArtifactLookupName = lookupName;
            AccountId = aId;
            ContactId = cId;
            IncidentId = iId;
        }

        /// <summary>
        /// This method retrieves artifact rules in a non-related record context with the necessary attributes to process an artifact rule
        /// </summary>
        /// <returns></returns>
        public List<mcs_artifactrule> GetArtifactRules()
        {
            using (var xrm = new CrmServiceContext(_orgService))
            {             
                return xrm.mcs_artifactruleSet
                        .Where(x => x.mcs_ParentRecord == _entityTarget.LogicalName && x.mcs_relatedRecord == null)
                        .Select(x => new mcs_artifactrule()
                        {
                            Id = x.Id,
                            mcs_QuestionIdentifier = x.mcs_QuestionIdentifier,
                            mcs_SuccessIndicator = x.mcs_SuccessIndicator,
                            mcs_name = x.mcs_name,
                            mcs_artifactruleId = x.mcs_artifactruleId,
                            mcs_IsMandatory = x.mcs_IsMandatory,
                            mcs_SpecifierLookup = x.mcs_SpecifierLookup,
                            mcs_Specifier = x.mcs_Specifier,
                            mcs_CommentsInstructions = x.mcs_CommentsInstructions,
                            mcs_ArtifactType = x.mcs_ArtifactType
                        }).ToList();
            }
        }

        /// <summary>
        /// This method retrieves artifact rules in an related record context with the necessary attributes to process an artifact rule
        /// </summary>
        /// <returns></returns>
        public List<mcs_artifactrule> GetArtifactRuleRelatedRecord()
        {
            using (var xrm = new CrmServiceContext(_orgService))
            {
                return xrm.mcs_artifactruleSet
                .Where(x => x.mcs_relatedRecord == _entityTarget.LogicalName)
                .Select(x => new mcs_artifactrule()
                {
                    mcs_QuestionIdentifier = x.mcs_QuestionIdentifier,
                    mcs_SuccessIndicator = x.mcs_SuccessIndicator,
                    mcs_name = x.mcs_name,
                    mcs_artifactruleId = x.mcs_artifactruleId,
                    mcs_IsMandatory = x.mcs_IsMandatory,                   
                    mcs_SpecifierLookup = x.mcs_SpecifierLookup,
                    mcs_Specifier = x.mcs_Specifier,
                    mcs_CommentsInstructions = x.mcs_CommentsInstructions,
                    mcs_ArtifactType = x.mcs_ArtifactType
                }).ToList();
            }
        }
        /// <summary>
        /// Creates an artifact for a given rule, ensuring the available attributes are populated on the artifact
        /// </summary>
        /// <param name="rule"></param>
        public void CreateArtifact(mcs_artifactrule rule)
        {
            var newArtifact = new mcs_artifact()
            {
                mcs_name = rule.mcs_name,
                mcs_ArtifactRuleId = rule.ToEntityReference(),
                [ArtifactLookupName] = _entityTarget.ToEntityReference(),
                mcs_ArtifactType = rule.mcs_ArtifactType,
                mcs_CommentsInstructions = rule.mcs_CommentsInstructions,
                mcs_ReviewStatus = new OptionSetValue(100000002) // Pending Review
            };

            // add portal Account, Contact, Incident lookups to Artifact
            if (AccountId != Guid.Empty) newArtifact.mcs_AccountId = new EntityReference("account", AccountId);
            if (ContactId != Guid.Empty) newArtifact.mcs_Contactid = new EntityReference("contact", ContactId);
            if (IncidentId != Guid.Empty) newArtifact.mcs_CaseId = new EntityReference("incident", IncidentId);

            // Set Association field based on Rule Specifier Lookup and or Specifier
            newArtifact.mcs_Association = GetArtifactAssociationByRule(rule);

            newArtifact.EntityState = EntityState.Created;
            _orgService.Create(newArtifact);

        }

        /// <summary>
        /// Conditionally checks specifier and specifier lookup fields to determine the association value for a given artifact rule and entity target
        /// </summary>
        /// <param name="rule"></param>
        /// <returns></returns>
        public string GetArtifactAssociationByRule(mcs_artifactrule rule)
        {
            string association = null;

            if (rule.mcs_SpecifierLookup == null && rule.mcs_Specifier != null)
            {
                association = (rule.mcs_Specifier.StartsWith("\"") && rule.mcs_Specifier.EndsWith("\"")) ?
                    rule.mcs_Specifier.Replace("\"", "") : (string)_entityTarget[rule.mcs_Specifier];
            }
            else if (rule.mcs_SpecifierLookup != null && _entityTarget.Contains(rule.mcs_SpecifierLookup))
            {
                // if lookup specified, retrieve string field for Association
                var lookup = (EntityReference)_entityTarget[rule.mcs_SpecifierLookup];
                var specifierEntity = _orgService.Retrieve(lookup.LogicalName, lookup.Id, new ColumnSet(new string[] { rule.mcs_Specifier }));
                association = (string)specifierEntity[rule.mcs_Specifier];
            }

            return association;
        }
        
        /// <summary>
        /// helper method to retrieve option set text
        /// </summary>
        /// <param name="entityName"></param>
        /// <param name="attributeName"></param>
        /// <param name="optionsetValue"></param>
        /// <param name="orgService"></param>
        /// <returns></returns>
        public static string GetOptionSetText(string entityName, string attributeName, int optionsetValue, IOrganizationService orgService)
        {          
            try
            {
               
                string optionsetText = string.Empty;
                RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest()
                {
                    EntityLogicalName = entityName,
                    LogicalName = attributeName,
                    RetrieveAsIfPublished = true
                };

              
                RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)orgService.Execute(retrieveAttributeRequest);
                PicklistAttributeMetadata picklistAttributeMetadata = (PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;

                OptionSetMetadata optionsetMetadata = picklistAttributeMetadata.OptionSet;

              
                foreach (OptionMetadata optionMetadata in optionsetMetadata.Options)
                {
                    if (optionMetadata.Value == optionsetValue)
                    {
                        optionsetText = optionMetadata.Label.UserLocalizedLabel.Label;
                        return optionsetText;
                    }

                }
                return optionsetText;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: Unable to load value for " + attributeName, ex);
            }
        }
    }
}
