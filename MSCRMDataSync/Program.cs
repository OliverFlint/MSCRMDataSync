using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
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
    class Program
    {
        private static string logfilename;

        static void Main(string[] args)
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

                var sourceType = sourceNode.Attributes["type"].Value;
                var sourceConnectionstring = sourceNode.InnerText;

                var destType = destNode.Attributes["type"].Value;
                var destConnectionstring = destNode.InnerText;

                var queryXml = queryNode.InnerText;

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
                    SaveToServer(sourceData, destConnectionstring);
                }
                else
                {
                    SaveToFile(sourceData, destConnectionstring);
                }

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
            var connection = CrmConnection.Parse(connectionstring);
            var service = new OrganizationService(connection);

            var query = new FetchExpression(queryxml);
            var results = service.RetrieveMultiple(query);
            return results.Entities.ToArray();
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
            var connection = CrmConnection.Parse(connectionstring);
            var service = new OrganizationService(connection);
            var batchSize = 10;
            var batchNo = 1;
            var currentIndex = 0;
            var moreBatches = true;

            var updatedEntities = new List<Entity>();
            var createEntities = entities.ToList();

            while (moreBatches)
            {
                var existsReq = new ExecuteMultipleRequest();
                existsReq.Requests = new OrganizationRequestCollection();
                existsReq.Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError=true,
                    ReturnResponses=true
                };

                while ((batchNo * batchSize) > currentIndex && currentIndex < entities.Length)
                {
                    var req = new RetrieveRequest();
                    req.ColumnSet = new ColumnSet();
                    req.Target = new EntityReference(entities[currentIndex].LogicalName, entities[currentIndex].Id);
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
                var createReq = new ExecuteMultipleRequest();
                createReq.Requests = new OrganizationRequestCollection();
                createReq.Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = true,
                    ReturnResponses = false
                };

                while ((createBatchNo * batchSize) > createIndex && createIndex < createEntities.Count)
                {
                    var req = new CreateRequest();
                    req.Target = createEntities[createIndex];
                    createReq.Requests.Add(req);
                    createIndex++;
                }

                service.Execute(createReq);

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
                var updateReq = new ExecuteMultipleRequest();
                updateReq.Requests = new OrganizationRequestCollection();
                updateReq.Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = true,
                    ReturnResponses = false
                };

                while ((updateBatchNo * batchSize) > updateIndex && updateIndex < updatedEntities.Count)
                {
                    var req = new UpdateRequest();
                    req.Target = updatedEntities[updateIndex];
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
