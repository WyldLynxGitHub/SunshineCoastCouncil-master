using HP.HPTRIM.Service.Client;
using HP.HPTRIM.ServiceModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestServicePost
{
    class Program
    {
        static void Main(string[] args)
        {
            TrimClient trimClient = new TrimClient("http://localhost/HPRMServiceAPI");
            trimClient.Credentials = System.Net.CredentialCache.DefaultCredentials;
            //
            //Locations request = new Locations();
            //request.q = "me";
            //request.Properties = new List<string>() { "SortName" };
            //Locations Loc = trimClient.Get<LocationsResponse>(request);

            //Console.WriteLine("Logged onto Eddie as " + Loc.Results[0].SortName);
            //
            //Properties = new PropertyList(PropertyIds.RecordTitle, PropertyIds.RecordNumber, PropertyIds.Url)
          
            Record record = new Record();
            record.Properties = new List<string>() { "RecordNumber", "Url" , "Container"};
            record.RecordType = new RecordTypeRef() { Uri = 2 };
            record.Title = "EPL test" + " - " + DateTime.Now.ToString();
            record.Classification = new ClassificationRef() { Uri = 3763 };
            record.Container = new RecordRef() { Uri = 6110 };
            record.Assignee = new LocationRef() { Uri = 5503 };
            //record.AssigneeStatus = new TrimProperty<RecLocSubTypes>() { }
                //RecLocSubTypes() { RecLocSubTypes.AtLocation };
            //RecLocSubTypes.AtLocation 
            //record.Assignee
            //record.SetCustomField("PropertyNumber", t1.PropertyNumber);
            //record.SetCustomField("ApplicationNumber", t1.ApplicationNumber);
            //record.SetCustomField("PropertyAddress", t1.PropertyAddress);
            RecordsResponse response = trimClient.Post<RecordsResponse>(record);
            if (response.Results.Count == 1)
            {
                Console.WriteLine("New property container created in Eddie - RM Uri: " + response.Results.First().Url + " " + response.Results.First().Container.Uri);
            }
            //
            //Records request = new Records()
            //{
            //    q = "all"
            //    ,
            //    Properties = new PropertyList(PropertyIds.RecordTitle, PropertyIds.RecordNumber, PropertyIds.Url)
            //};

            //RecordsResponse response = trimClient.Get<RecordsResponse>(request);

            //foreach (Record record in response.Results)
            //{
            //    Console.WriteLine(record.Title + " " + record.Number + " " + record.Url);
            //}
        }
    }
}
