# Dynamics 365 Artifact Management

The Artifact Management Solution is an add-on feature to Dynamics 365 Customer Engagement which acts as a rule engine for generating **_artifacts_**; a custom entity that hosts files uploaded by system users and portal users populated with contextual metadata. This solution benefits organizations with a need to require/recommend certain files be uploaded based on a defined set of rules.  It also benefits organizations with a desire to manage files solely in Dynamics 365 instead of leveraging SharePoint Integrations.

## Installing

For the latest stable version:

```bash
Import _ArtifactManagement_1_0_0_1.zip from the Releases folder into the desired Dynamics 365 CE Organization
```

## Configuration

Please refer to the [Wiki](https://github.com/brendon-colburn/d365-artifact-management/wiki) for details regarding configuration of the Artifact Management Solution

## Contribute

There are many ways to contribute to Artifact Management.

* [Submit bugs](https://github.com/brendon-colburn/d365-artifact-management/issues) and help us verify fixes as they are checked in.
* **Call to action:** Form JavaScript/TypeScript solution that populates the artifact rule form fields with a dynamic picklist of appropriate options.
* **Call to action:** A solution that automatically creates the necessary relationships and sdk message processing steps for a new entity (*see [Configuration for New Entities on the Wiki](https://github.com/brendon-colburn/d365-artifact-management/wiki/Configuration-for-New-Entities)*)
* Last but not least, please feel free to fork this for your own solutions.  As you come up with improvements reach out and we may add you as a contributor and merge parts of your branch to master.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see 
the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) 
with any additional questions or comments.