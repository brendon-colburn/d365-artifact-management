using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using Microsoft.Xrm.Sdk.Query;

namespace MCS.ArtifactManagement
{
    /// <summary>
    /// Workflow
    /// Registered against Annotation On Create
    /// 
    /// Update parent Artifact Entity Flag and Date
    /// 
    /// </summary>
    public class ArtifactUploadCheck : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);                     

            try
            {               
                var targetId = context.PrimaryEntityId;
                var targetName = context.PrimaryEntityName;               
               
                var thisAnnotation = service.Retrieve(targetName, targetId, new ColumnSet(new string[] {"objectid","createdon","notetext"}));

                var objectRef = (EntityReference)thisAnnotation["objectid"];
                
                // If this annotation is for an artifact update the Artifact
                if (objectRef.LogicalName == "mcs_artifact")
                {
                    var updateArtifact = new mcs_artifact()
                    {
                        Id = objectRef.Id,
                        mcs_Upload = new OptionSetValue(1),  // Uploaded
                        mcs_UploadDate = (DateTime)thisAnnotation["createdon"]
                    };

                    service.Update(updateArtifact);
                    
                    // Update the Annotation if it doesn't have *WEB* in the notetext field 
                    // This is required for portal visibility

                    if (!thisAnnotation.Contains("notetext"))
                    {
                        thisAnnotation["notetext"] = "*WEB*";
                        service.Update(thisAnnotation);
                    }
                    else if (((string)thisAnnotation["notetext"]).IndexOf("*WEB*") == -1)
                    {
                        thisAnnotation["notetext"] = (string)thisAnnotation["notetext"] + "*WEB*";
                        service.Update(thisAnnotation);
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