using Microsoft.Crm.Sdk.Messages;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Text.RegularExpressions;

namespace ModelDrivenAppCloner
{
    class Program
    {
        // Global service object for interacting with Dataverse
        static IOrganizationService _service;

        // Enum to represent different component types in Dataverse
        enum ComponentType
        {
            Entity = 1,
            Attribute = 2,
            Relationship = 3,
            AttributePicklistValue = 4,
            AttributeLookupValue = 5,
            ViewAttribute = 6,
            LocalizedLabel = 7,
            RelationshipExtraCondition = 8,
            OptionSet = 9,
            EntityRelationship = 10,
            EntityRelationshipRole = 11,
            EntityRelationshipRelationships = 12,
            ManagedProperty = 13,
            EntityKey = 14,
            Role = 20,
            RolePrivilege = 21,
            DisplayString = 22,
            DisplayStringMap = 23,
            Form = 24,
            Organization = 25,
            SavedQuery = 26,
            Workflow = 29,
            Report = 31,
            ReportEntity = 32,
            ReportCategory = 33,
            ReportVisibility = 34,
            Attachment = 35,
            EmailTemplate = 36,
            ContractTemplate = 37,
            KBArticleTemplate = 38,
            MailMergeTemplate = 39,
            DuplicateRule = 44,
            DuplicateRuleCondition = 45,
            EntityMap = 46,
            AttributeMap = 47,
            RibbonCommand = 48,
            RibbonContextGroup = 49,
            RibbonCustomization = 50,
            RibbonRule = 52,
            RibbonTabToCommandMap = 53,
            RibbonDiff = 55,
            SavedQueryVisualization = 59,
            SystemForm = 60,
            WebResource = 61,
            SiteMap = 62,
            ConnectionRole = 63,
            HierarchyRule = 65,
            CustomControl = 66,
            CustomControlDefaultConfig = 68,
            FieldSecurityProfile = 70,
            FieldPermission = 71,
            PluginType = 90,
            PluginAssembly = 91,
            SDKMessageProcessingStep = 92,
            SDKMessageProcessingStepImage = 93,
            ServiceEndpoint = 95,
            RoutingRule = 150,
            RoutingRuleItem = 151,
            SLA = 152,
            SLAItem = 153,
            ConvertRule = 154,
            ConvertRuleItem = 155,
            MobileOfflineProfile = 161,
            MobileOfflineProfileItem = 162,
            SimilarityRule = 165
        }

        // Dictionary to map ComponentType to logical names
        static Dictionary<ComponentType, string> ComponentTypeToLogicalName = new()
        {
            { ComponentType.Entity, "entity" },
            { ComponentType.Attribute, "attribute" },
            { ComponentType.Relationship, "relationship" },
            { ComponentType.OptionSet, "optionset" },
            { ComponentType.EntityKey, "entitykey" },
            { ComponentType.Role, "role" },
            { ComponentType.Form, "systemform" },
            { ComponentType.SavedQuery, "savedquery" },
            { ComponentType.Workflow, "workflow" },
            { ComponentType.Report, "report" },
            { ComponentType.EmailTemplate, "template" },
            { ComponentType.ContractTemplate, "contracttemplate" },
            { ComponentType.KBArticleTemplate, "kbarticletemplate" },
            { ComponentType.MailMergeTemplate, "mailmergetemplate" },
            { ComponentType.DuplicateRule, "duplicaterule" },
            { ComponentType.RibbonCustomization, "ribboncustomization" },
            { ComponentType.SavedQueryVisualization, "savedqueryvisualization" },
            { ComponentType.SystemForm, "systemform" },
            { ComponentType.WebResource, "webresource" },
            { ComponentType.SiteMap, "sitemap" },
            { ComponentType.ConnectionRole, "connectionrole" },
            { ComponentType.HierarchyRule, "hierarchyrule" },
            { ComponentType.CustomControl, "customcontrol" },
            { ComponentType.FieldSecurityProfile, "fieldsecurityprofile" },
            { ComponentType.PluginType, "plugintype" },
            { ComponentType.PluginAssembly, "pluginassembly" },
            { ComponentType.SDKMessageProcessingStep, "sdkmessageprocessingstep" },
            { ComponentType.SDKMessageProcessingStepImage, "sdkmessageprocessingstepimage" },
            { ComponentType.ServiceEndpoint, "serviceendpoint" },
            { ComponentType.RoutingRule, "routingrule" },
            { ComponentType.SLA, "sla" },
            { ComponentType.ConvertRule, "convertrule" },
            { ComponentType.MobileOfflineProfile, "mobileofflineprofile" },
            { ComponentType.SimilarityRule, "similarityrule" }
        };

        static void Main(string[] args)
        {
            try
            {
                ConnectToDataverse();

                Console.Write("Enter the name of the source app: ");
                string sourceAppName = Console.ReadLine();
                Guid sourceAppId = GetAppIdByName(sourceAppName);

                Console.Write("Enter the name for the cloned app: ");
                string clonedAppName = Console.ReadLine();

                Console.Write("Enter the name for the new solution: ");
                string newSolutionName = Console.ReadLine();

                string newSolutionUniqueName = CreateSolution(newSolutionName);

                Guid newAppId = CloneModelDrivenApp(sourceAppId, newSolutionUniqueName, clonedAppName);
                PublishApp(newAppId);

                Console.WriteLine("App cloned and published successfully!");

                string appInfo = GetAppUrl(newAppId);
                Console.WriteLine("\nNew App Information:");
                Console.WriteLine(appInfo);
            }
            catch (System.ServiceModel.FaultException<OrganizationServiceFault> ex)
            {
                Console.WriteLine($"Error code: {ex.Detail.ErrorCode}");
                Console.WriteLine($"Error message: {ex.Detail.Message}");
                Console.WriteLine($"Error details: {ex.Detail.TraceText}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// Connects to Dataverse using the provided connection string.
        /// </summary>
        static void ConnectToDataverse()
        {
            // TODO: Replace these values with your actual Dataverse environment details
            string clientId = "";
            string clientSecret = "";
            string environment = "";
            var connectionString = @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";
            ServiceClient crmServiceClient = new(connectionString);
            _service = crmServiceClient;
        }

        /// <summary>
        /// Retrieves the ID of an app by its name.
        /// </summary>
        /// <param name="appName">The name of the app to find.</param>
        /// <returns>The GUID of the app.</returns>
        static Guid GetAppIdByName(string appName)
        {
            QueryExpression query = new("appmodule")
            {
                ColumnSet = new ColumnSet("appmoduleid")
            };
            query.Criteria.AddCondition("name", ConditionOperator.Equal, appName);

            EntityCollection result = _service.RetrieveMultiple(query);

            if (result.Entities.Count == 0)
            {
                throw new Exception($"No app found with name: {appName}");
            }

            return result.Entities[0].Id;
        }

        /// <summary>
        /// Retrieves the ID of a publisher by its name.
        /// </summary>
        /// <param name="publisherName">The name of the publisher to find.</param>
        /// <returns>The GUID of the publisher.</returns>
        static Guid GetPublisherIdByName(string publisherName)
        {
            QueryExpression query = new("publisher")
            {
                ColumnSet = new ColumnSet("publisherid")
            };
            query.Criteria.AddCondition("uniquename", ConditionOperator.Equal, publisherName);

            EntityCollection result = _service.RetrieveMultiple(query);

            if (result.Entities.Count == 0)
            {
                throw new Exception($"No publisher found with name: {publisherName}");
            }

            return result.Entities[0].Id;
        }

        /// <summary>
        /// Creates a new solution in Dataverse.
        /// </summary>
        /// <param name="solutionName">The name of the solution to create.</param>
        /// <returns>The unique name of the created solution.</returns>
        static string CreateSolution(string solutionName)
        {
            string uniqueName = new string(solutionName.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray()).ToLower();
            uniqueName = $"customsolution_{uniqueName}";

            Entity solution = new("solution");
            solution["uniquename"] = uniqueName;
            solution["friendlyname"] = solutionName;
            solution["version"] = "1.0.0.0";
            Guid publisherId = GetPublisherIdByName("ibm");
            solution["publisherid"] = new EntityReference("publisher", publisherId);

            Guid solutionId = _service.Create(solution);
            Console.WriteLine($"New solution created with ID: {solutionId}");

            return uniqueName;
        }

        /// <summary>
        /// Creates a copy of a sitemap for the new app.
        /// </summary>
        /// <param name="sourceSiteMapId">The ID of the source sitemap.</param>
        /// <param name="targetSolutionName">The name of the target solution.</param>
        /// <param name="newAppName">The name of the new app.</param>
        /// <returns>The GUID of the new sitemap.</returns>
        static Guid CreateSiteMapCopy(Guid sourceSiteMapId, string targetSolutionName, string newAppName)
        {
            Entity sourceSiteMap = _service.Retrieve("sitemap", sourceSiteMapId, new ColumnSet(true));

            Entity newSiteMap = new("sitemap");
            newSiteMap["sitemapname"] = newAppName;

            string sanitizedName = Regex.Replace(newAppName, @"[^a-zA-Z0-9_]", "");

            string uniqueName = $"ibm_{sanitizedName}";
            if (!char.IsLetter(uniqueName[0]))
            {
                uniqueName = "sitemap_" + uniqueName;
            }

            if (uniqueName.Length > 100)
            {
                uniqueName = uniqueName.Substring(0, 100);
            }

            newSiteMap["sitemapnameunique"] = uniqueName;
            newSiteMap["sitemapxml"] = sourceSiteMap["sitemapxml"];
            newSiteMap["isappaware"] = true;

            CreateRequest createRequest = new()
            {
                Target = newSiteMap
            };
            createRequest.Parameters["SolutionUniqueName"] = targetSolutionName;

            try
            {
                CreateResponse createResponse = (CreateResponse)_service.Execute(createRequest);
                Console.WriteLine($"New sitemap created with ID: {createResponse.id}");
                Console.WriteLine($"Sitemap unique name: {uniqueName}");
                return createResponse.id;
            }
            catch (System.ServiceModel.FaultException<OrganizationServiceFault> ex)
            {
                Console.WriteLine($"Error creating sitemap:");
                Console.WriteLine($"Error code: {ex.Detail.ErrorCode}");
                Console.WriteLine($"Error message: {ex.Detail.Message}");
                Console.WriteLine($"Trace: {ex.Detail.TraceText}");
                throw;
            }
        }

        /// <summary>
        /// Clones a model-driven app.
        /// </summary>
        /// <param name="sourceAppId">The ID of the source app to clone.</param>
        /// <param name="targetSolutionName">The name of the target solution.</param>
        /// <param name="clonedAppName">The name for the cloned app.</param>
        /// <returns>The GUID of the new app.</returns>
        static Guid CloneModelDrivenApp(Guid sourceAppId, string targetSolutionName, string clonedAppName)
        {
            Entity sourceApp = _service.Retrieve("appmodule", sourceAppId, new ColumnSet(true));

            try
            {
                Entity newApp = new("appmodule");
                newApp["name"] = clonedAppName;

                string sanitizedName = Regex.Replace(clonedAppName, @"[^a-zA-Z0-9_]", "");

                string uniqueName = $"ibm_{sanitizedName}";
                if (!char.IsLetter(uniqueName[0]))
                {
                    uniqueName = "app_" + uniqueName;
                }

                if (uniqueName.Length > 100)
                {
                    uniqueName = uniqueName.Substring(0, 100);
                }

                newApp["uniquename"] = uniqueName;

                // Copy relevant attributes from the source app
                if (sourceApp.Contains("description"))
                    newApp["description"] = sourceApp["description"];
                if (sourceApp.Contains("webresourceid"))
                    newApp["webresourceid"] = sourceApp["webresourceid"];
                if (sourceApp.Contains("clienttype"))
                    newApp["clienttype"] = sourceApp["clienttype"];
                if (sourceApp.Contains("navigationtype"))
                    newApp["navigationtype"] = sourceApp["navigationtype"];
                if (sourceApp.Contains("formfactor"))
                    newApp["formfactor"] = sourceApp["formfactor"];

                newApp["isfeatured"] = false;

                Guid newAppId = _service.Create(newApp);
                Console.WriteLine($"New app created with ID: {newAppId}");
                Console.WriteLine($"Unique name: {uniqueName}");

                // Retrieve components from the source app
                RetrieveAppComponentsRequest componentsRequest = new()
                {
                    AppModuleId = sourceAppId
                };
                RetrieveAppComponentsResponse componentsResponse = (RetrieveAppComponentsResponse)_service.Execute(componentsRequest);

                EntityReferenceCollection newComponents = [];

                // Process each component
                foreach (var component in componentsResponse.AppComponents.Entities)
                {
                    ComponentType componentType = (ComponentType)component.GetAttributeValue<OptionSetValue>("componenttype").Value;
                    Guid objectId = component.GetAttributeValue<Guid>("objectid");

                    if (ComponentTypeToLogicalName.TryGetValue(componentType, out string logicalName))
                    {
                        if (componentType == ComponentType.SiteMap)
                        {
                            Guid newSiteMapId = CreateSiteMapCopy(objectId, targetSolutionName, clonedAppName);
                            newComponents.Add(new EntityReference(logicalName, newSiteMapId));
                        }
                        else if (componentType == ComponentType.Entity)
                        {
                            var entityMetadata = ((RetrieveEntityResponse)_service.Execute(new RetrieveEntityRequest { MetadataId = objectId })).EntityMetadata;
                            newComponents.Add(new EntityReference(entityMetadata.LogicalName, objectId));
                        }
                        else
                        {
                            newComponents.Add(new EntityReference(logicalName, objectId));
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Component type {componentType} not handled.");
                    }
                }

                // Add components to the new app
                AddAppComponentsRequest addComponentsRequest = new()
                {
                    AppId = newAppId,
                    Components = newComponents
                };
                _service.Execute(addComponentsRequest);

                // Add the new app to the solution
                AddSolutionComponentRequest addToSolutionRequest = new()
                {
                    AddRequiredComponents = false,
                    ComponentId = newAppId,
                    ComponentType = 80, // 80 represents App Module
                    SolutionUniqueName = targetSolutionName
                };
                _service.Execute(addToSolutionRequest);

                return newAppId;
            }
            catch (System.ServiceModel.FaultException<OrganizationServiceFault> ex)
            {
                Console.WriteLine($"Error code: {ex.Detail.ErrorCode}");
                Console.WriteLine($"Error message: {ex.Detail.Message}");
                Console.WriteLine($"Trace: {ex.Detail.TraceText}");
                throw;
            }
        }

        /// <summary>
        /// Publishes the newly created app.
        /// </summary>
        /// <param name="appId">The ID of the app to publish.</param>
        static void PublishApp(Guid appId)
        {
            Console.WriteLine("Publishing the app...");
            PublishXmlRequest publishRequest = new()
            {
                ParameterXml = $"<importexportxml><appmodules><appmodule>{appId}</appmodule></appmodules></importexportxml>"
            };
            _service.Execute(publishRequest);
            Console.WriteLine("App published successfully.");
        }

        /// <summary>
        /// Retrieves the URL for the newly created app.
        /// </summary>
        /// <param name="appId">The ID of the app.</param>
        /// <returns>A string containing app information and URL.</returns>
        static string GetAppUrl(Guid appId)
        {
            try
            {
                // First, get the app details
                var query = new QueryExpression("appmodule")
                {
                    ColumnSet = new ColumnSet("name", "uniquename"),
                    Criteria = new FilterExpression
                    {
                        Conditions = {
                            new ConditionExpression("appmoduleid", ConditionOperator.Equal, appId)
                        }
                    }
                };

                var result = _service.RetrieveMultiple(query);

                if (result.Entities.Count > 0)
                {
                    var app = result.Entities[0];
                    string appName = app.GetAttributeValue<string>("name");
                    string appUniqueName = app.GetAttributeValue<string>("uniquename");

                    // Get the organization URL
                    var request = new RetrieveCurrentOrganizationRequest();
                    var response = (RetrieveCurrentOrganizationResponse)_service.Execute(request);

                    if (response.Detail != null)
                    {
                        // Use the correct EndpointType
                        string orgUrl = response.Detail.Endpoints[Microsoft.Xrm.Sdk.Organization.EndpointType.WebApplication];

                        if (!string.IsNullOrEmpty(orgUrl))
                        {
                            // Ensure the URL ends with a slash
                            if (!orgUrl.EndsWith("/"))
                            {
                                orgUrl += "/";
                            }

                            // Construct the full URL
                            string fullUrl = $"{orgUrl}main.aspx?appid={appId}";

                            return $"App Name: {appName}\nApp Unique Name: {appUniqueName}\nApp URL: {fullUrl}";
                        }
                        else
                        {
                            return "Organization URL not found.";
                        }
                    }
                    else
                    {
                        return "Organization details not found.";
                    }
                }
                else
                {
                    return "App not found.";
                }
            }
            catch (Exception ex)
            {
                return $"Error retrieving app URL: {ex.Message}";
            }
        }
    }
}