using Microsoft.Xrm.Sdk;
using System;
using System.Linq;

namespace MCS.ArtifactManagement
{
    /// <summary>
    /// Plugin ProcessingDelete  
    /// registered against  incident, gnext_expense, gnext_coverage
    /// preValidate - Delete Stage only
    /// 
    /// Add Plugin Unsecure Configuration data in this format
    /// Artifact Lookup field name for the entity  ex.  mcs_caseid
    /// 
    /// When an entity is deleted that could have child Artifacts records those child artifacts are deleted.
    /// Since there are 3 separate parent entities for the artifact they can not all have parental deletion
    /// </summary>
    public class ProcessingDelete : IPlugin
    {        
        private string _secureConfig = null;
        private string _unsecureConfig = null;

        public ProcessingDelete(string unsecureConfig, string secureConfig)
        {
            _secureConfig = secureConfig;
            _unsecureConfig = unsecureConfig;
        }        
           
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            try
            {
                // The lookup field name in the mcs_artifact entity for the target parent entity that is being deleted
                var targetEntityArtifactLookupFieldName = _unsecureConfig.Trim();               

                using (var xrm = new CrmServiceContext(service))
                {
                    var relatedArtifacts = xrm.mcs_artifactSet.Where(x => ((EntityReference)x[targetEntityArtifactLookupFieldName]).Id == context.PrimaryEntityId)
                    .Select(x => new mcs_artifact()
                    {
                        Id = x.Id                        
                    });

                    relatedArtifacts.ToList().ForEach(z => service.Delete("mcs_artifact", z.Id));
                }

            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

    }
}