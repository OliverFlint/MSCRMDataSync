using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace MSCRMDataSync
{
    public class Program
    {
        private static string logfilename;
        private static int batchSize = 10;

        public static void Main(string[] args)
        {
            try
            {
                logfilename = string.Format("{0}.log", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);
                log("Start");
                if (args.Length != 1)
                {
                    throw new ArgumentException("Config file argument missing!");
                }

                var config = new XmlDocument();
                config.Load(args[0]);

                var sourceNode = config.SelectSingleNode("//mscrmdatasync/source");
                var destNode = config.SelectSingleNode("//mscrmdatasync/destination");
                var queryNode = config.SelectSingleNode("//mscrmdatasync/query");
                var batchsizeNode = config.SelectSingleNode("//mscrmdatasync/batchsize");
                var syncTypeNode = config.SelectSingleNode("//mscrmdatasync/type");

                var sourceType = sourceNode.Attributes["type"].Value;
                var sourceConnectionstring = sourceNode.InnerText;

                var destType = destNode.Attributes["type"].Value;
                var destConnectionstring = destNode.InnerText;

                var syncType = syncTypeNode != null ? syncTypeNode.InnerText : String.Empty;

                var queryXml = queryNode.InnerText;

                batchSize = int.Parse(batchsizeNode.InnerText);

                Entity[] sourceData = null;

                if (sourceType.ToLower() == "server")
                {
                    sourceData = GetEntitiesFromServer(sourceConnectionstring, queryXml);
                }
                else
                {
                    sourceData = GetEntitiesFromFile(sourceConnectionstring);
                }

                if (destType.ToLower() == "server")
                {
                    if (syncType.ToLower() == "manytomany")
                    {
                        SaveN2NToServerSingle(sourceData, destConnectionstring);
                    }
                    else
                    {
                        SaveToServer(sourceData, destConnectionstring);
                    }
                }
                else
                {
                    SaveToFile(sourceData, destConnectionstring);
                }
                logsuccess("Job completed successfully");
            }
            catch (Exception ex)
            {
                logerror(ex.Message);
            }
            finally
            {
                log("End");
            }
        }

        private static Entity[] GetEntitiesFromServer(string connectionstring, string queryxml)
        {
            var service = new Microsoft.Xrm.Tooling.Connector.CrmServiceClient(connectionstring);

            var resp = (FetchXmlToQueryExpressionResponse)service.Execute(new FetchXmlToQueryExpressionRequest() { FetchXml=queryxml });
            var query = resp.Query;
            int page = 1;

            var results = ExecutePagedQuery(service, query, page);
            var resultList = results.Entities.ToList();
            while (results.MoreRecords)
            {
                page++;
                results = ExecutePagedQuery(service, query, page, results.PagingCookie);
                resultList.AddRange(results.Entities.ToList());
            }

            return resultList.ToArray();
        }

        private static EntityCollection ExecutePagedQuery(Microsoft.Xrm.Tooling.Connector.CrmServiceClient service, QueryExpression query, int page, string pagingCookie = null)
        {
            query.PageInfo = new PagingInfo() { Count=1000, PageNumber=page };
            if (pagingCookie != null) query.PageInfo.PagingCookie = pagingCookie;
            return service.RetrieveMultiple(query);
        }

        private static Entity[] GetEntitiesFromFile(string filename)
        {
            return DataContractDeSerializeObject(filename);
        }

        private static void SaveToFile(Entity[] entities, string filename)
        {
            var entitiesxml = DataContractSerializeObject<Entity[]>(entities);
            File.WriteAllText(filename, entitiesxml);
        }

        private static void SaveToServer(Entity[] entities, string connectionstring)
        {
            if (batchSize > 1)
            {
                SaveToServerBatched(entities, connectionstring);
            }
            else
            {
                SaveToServerSingle(entities, connectionstring);
            }
        }

        private static void SaveToServerSingle(Entity[] entities, string connectionstring)
        {
            var service = new Microsoft.Xrm.Tooling.Connector.CrmServiceClient(connectionstring);

            foreach (var e in entities)
            {
                var exists = false;
                try
                {
                    var existingEntity = service.Retrieve(e.LogicalName, e.Id, new ColumnSet());
                    if (existingEntity != null)
                    {
                        exists = true;
                    }
                }
                catch { }

                if (exists)
                {
                    service.Update(e);
                }
                else
                {
                    service.Create(e);
                }
            }
        }

        private static void SaveN2NToServerSingle(Entity[] entities, string connectionstring)
        {
            var service = new Microsoft.Xrm.Tooling.Connector.CrmServiceClient(connectionstring);

            var req = new RetrieveRelationshipRequest
            {
                Name = entities[0].LogicalName
            };
            var metadata = (RetrieveRelationshipResponse)service.Execute(req);
            var relationship = (ManyToManyRelationshipMetadata)metadata.RelationshipMetadata;

            foreach (var e in entities)
            {
                var exists = false;
                try
                {
                    var existingEntity = service.Retrieve(e.LogicalName, e.Id, new ColumnSet());
                    if (existingEntity != null)
                    {
                        exists = true;
                    }
                }
                catch { }

                if (!exists)
                {
                    var rc = new EntityReferenceCollection();
                    rc.Add(new EntityReference(relationship.Entity2LogicalName, e.GetAttributeValue<Guid>(relationship.Entity2IntersectAttribute)));
                    service.Associate(relationship.Entity1LogicalName
                        , e.GetAttributeValue<Guid>(relationship.Entity1IntersectAttribute)
                        , new Relationship(relationship.SchemaName)
                        , rc);
                }
            }
        }

        private static void SaveToServerBatched(Entity[] entities, string connectionstring)
        {
            var service = new Microsoft.Xrm.Tooling.Connector.CrmServiceClient(connectionstring);

            //var batchSize = 10;
            var batchNo = 1;
            var currentIndex = 0;
            var moreBatches = true;

            var updatedEntities = new List<Entity>();
            var createEntities = entities.ToList();

            while (moreBatches)
            {
                var existsReq = new ExecuteMultipleRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = true,
                        ReturnResponses = true
                    }
                };

                while ((batchNo * batchSize) > currentIndex && currentIndex < entities.Length)
                {
                    var req = new RetrieveRequest
                    {
                        ColumnSet = new ColumnSet(),
                        Target = new EntityReference(entities[currentIndex].LogicalName, entities[currentIndex].Id)
                    };
                    existsReq.Requests.Add(req);
                    currentIndex++;
                }

                var existsResp = (ExecuteMultipleResponse)service.Execute(existsReq);

                //Add updated entities to the updatedEntities collection
                foreach (var item in existsResp.Responses)
                {
                    if (item.Response != null)
                    {
                        var updatedentity = entities.First(e => e.Id == ((Entity)item.Response.Results.FirstOrDefault().Value).Id);
                        updatedEntities.Add(updatedentity);

                        //Remove the entity form the createdEntities collection
                        createEntities.Remove(updatedentity);
                    }
                }

                if (currentIndex >= entities.Length)
                {
                    moreBatches = false;
                }
                else
                {
                    moreBatches = true;
                    batchNo++;
                }
            }
            
            //Create entities
            var createIndex = 0;
            var createBatchNo = 1;
            moreBatches = true;
            while (moreBatches)
            {
                var createReq = new ExecuteMultipleRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = true,
                        ReturnResponses = false
                    }
                };

                while ((createBatchNo * batchSize) > createIndex && createIndex < createEntities.Count)
                {
                    var req = new CreateRequest
                    {
                        Target = createEntities[createIndex]
                    };
                    createReq.Requests.Add(req);
                    createIndex++;
                }

                var createResp = (ExecuteMultipleResponse)service.Execute(createReq);

                if(createResp.IsFaulted){
                    foreach (var s in createResp.Responses)
                    {
                        if (s.Fault != null)
                        {
                            logerror(s.Fault.Message);
                        }
                    }
                }

                if (createIndex >= createEntities.Count)
                {
                    moreBatches = false;
                }
                else
                {
                    moreBatches = true;
                    createBatchNo++;
                }
            }

            //Update Entities
            var updateIndex = 0;
            var updateBatchNo = 1;
            moreBatches = true;
            while (moreBatches)
            {
                var updateReq = new ExecuteMultipleRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = true,
                        ReturnResponses = false
                    }
                };

                while ((updateBatchNo * batchSize) > updateIndex && updateIndex < updatedEntities.Count)
                {
                    var req = new UpdateRequest
                    {
                        Target = updatedEntities[updateIndex]
                    };
                    updateReq.Requests.Add(req);
                    updateIndex++;
                }

                service.Execute(updateReq);

                if (updateIndex >= updatedEntities.Count)
                {
                    moreBatches = false;
                }
                else
                {
                    moreBatches = true;
                    updateBatchNo++;
                }
            }
        }

        private static string DataContractSerializeObject<T>(T objectToSerialize)
        {
            using (MemoryStream memStm = new MemoryStream())
            {
                var serializer = new DataContractSerializer(typeof(T));
                serializer.WriteObject(memStm, objectToSerialize);

                memStm.Seek(0, SeekOrigin.Begin);

                using (var streamReader = new StreamReader(memStm))
                {
                    string result = streamReader.ReadToEnd();
                    return result;
                }
            }
        }

        private static Entity[] DataContractDeSerializeObject(string filename)
        {
            using (var file = new FileStream(filename, FileMode.Open))
            {
                var serializer = new DataContractSerializer(typeof(Entity[]));
                var entities = (Entity[])serializer.ReadObject(file);
                return entities;
            }
        }

        private static void logerror(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(message);
            Console.ResetColor();
            logtofile(string.Format("{0} - {1}", "ERROR", message));
        }

        private static void logsuccess(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(message);
            Console.ResetColor();
            logtofile(string.Format("{0} - {1}", "SUCCESS", message));
        }

        private static void log(string message)
        {
            Console.WriteLine(message);
            logtofile(string.Format("{0} - {1}", "INFO", message));
        }

        private static void logtofile(string message)
        {
            try
            {
                File.AppendAllText(logfilename, string.Format("{0}: {1}{2}", DateTime.Now, message, Environment.NewLine));
            }
            catch { }
        }
    }
}
