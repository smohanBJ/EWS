using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Management;
using System.Net;

namespace HelloWorld
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
            // service.UseDefaultCredentials = true;
            //service.Credentials = new WebCredentials(CredentialCache.DefaultNetworkCredentials);
            //service.TraceEnabled = true;

            service.Credentials = new WebCredentials("siddharth.1@exchange.bluejeansdev.com", "denim@123");
            service.Url = new Uri("https://exchange-server.exchange.bluejeansdev.com/EWS/Exchange.asmx");
            
            //Certification bypass
            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;

            DateTime startDate = DateTime.Now;
            DateTime endDate = startDate.AddDays(30);
            //Specify number of meetings to see
            const int NUM_APPTS = 5;

            /* Code to show all calendars of a user */
            Folder rootfolder = Folder.Bind(service, WellKnownFolderName.Calendar);

             Console.WriteLine("The " + rootfolder.DisplayName + " has " + rootfolder.ChildFolderCount + " child folders.");
             // A GetFolder operation has been performed.
             // Now do something with the folder, such as display each child folder's name and ID.
             rootfolder.Load();
             foreach (Folder folder in rootfolder.FindFolders(new FolderView(100)))
             {
                 Console.WriteLine("\nName: " + folder.DisplayName + "\n  Id: " + folder.Id);
                 // Initialize the calendar folder object with only the folder ID. 
                 CalendarFolder calendar = CalendarFolder.Bind(service, folder.Id, new PropertySet());

                 // Set the start and end time and number of appointments to retrieve.
                 CalendarView cView = new CalendarView(startDate, endDate, NUM_APPTS);

                 // Limit the properties returned to the appointment's subject, start time, and end time.
                 cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End);

                 // Retrieve a collection of appointments by using the calendar view.
                 FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

                 Console.WriteLine("\nThe first " + NUM_APPTS + " appointments on your calendar from " + startDate.Date.ToShortDateString() +
                                   " to " + endDate.Date.ToShortDateString() + " are: \n");
                 Console.WriteLine(appointments.TotalCount);
                 foreach (Appointment a in appointments)
                 {
                     Console.Write("Subject: " + a.Subject.ToString() + " ");
                     Console.Write("Start: " + a.Start.ToString() + " ");
                     Console.Write("End: " + a.End.ToString());
                     Console.WriteLine();
                 }
             }
             
            
             Console.WriteLine("Get all shared calendars");

            /*Code to get meetings from all rooms*/
             Dictionary<string, string>  result = GetSharedCalendarFolders(service, "siddharth.1@exchange.bluejeansdev.com");
              foreach (KeyValuePair<string, string> kvp in result)
              {
                  Console.WriteLine("Key = {0}, Value = {1}", kvp.Key, kvp.Value);

                FolderId te = new FolderId(WellKnownFolderName.Calendar, kvp.Value);

                DateTime start = DateTime.Now;
                DateTime end = DateTime.Now.AddDays(30);

                CalendarView view = new CalendarView(start, end);

                foreach (Appointment a in service.FindAppointments(te, view))
                {
                    Console.Write("Subject: " + a.Subject.ToString() + " ");
                    Console.Write("Start: " + a.Start.ToString() + " ");
                    Console.Write("End: " + a.End.ToString());
                    Console.WriteLine();
                }
            }
              

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();

        }

        static Dictionary<string, string> GetSharedCalendarFolders(ExchangeService service, String mbMailboxname)
        {
            Dictionary<String, String> rtList = new System.Collections.Generic.Dictionary<string, String>();

            FolderId rfRootFolderid = new FolderId(WellKnownFolderName.Root, mbMailboxname);
            FolderView fvFolderView = new FolderView(int.MaxValue);
            SearchFilter sfSearchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Common Views");
            FindFoldersResults ffoldres = service.FindFolders(rfRootFolderid, sfSearchFilter, fvFolderView);
            if (ffoldres.Folders.Count >= 1)
            {

                PropertySet psPropset = new PropertySet(BasePropertySet.FirstClassProperties);
                ExtendedPropertyDefinition PidTagWlinkAddressBookEID = new ExtendedPropertyDefinition(0x6854, MapiPropertyType.Binary);
                ExtendedPropertyDefinition PidTagWlinkGroupName = new ExtendedPropertyDefinition(0x6851, MapiPropertyType.String);

                psPropset.Add(PidTagWlinkAddressBookEID);
                ItemView iv = new ItemView(1000);
                iv.PropertySet = psPropset;
                iv.Traversal = ItemTraversal.Associated;

                SearchFilter cntSearch = new SearchFilter.IsEqualTo(PidTagWlinkGroupName, "People's Calendars");
                // Can also find this using PidTagWlinkType = wblSharedFolder
                FindItemsResults<Item> fiResults = ffoldres.Folders[0].FindItems(cntSearch, iv);
                foreach (Item itItem in fiResults.Items)
                {
                    try
                    {
                        object GroupName = null;
                        object WlinkAddressBookEID = null;

                        if (itItem.TryGetProperty(PidTagWlinkAddressBookEID, out WlinkAddressBookEID))
                        {

                            byte[] ssStoreID = (byte[])WlinkAddressBookEID;
                            int leLegDnStart = 0;
                            
                            String lnLegDN = "";
                            for (int ssArraynum = (ssStoreID.Length - 2); ssArraynum != 0; ssArraynum--)
                            {
                                if (ssStoreID[ssArraynum] == 0)
                                {
                                    leLegDnStart = ssArraynum;
                                    lnLegDN = System.Text.ASCIIEncoding.ASCII.GetString(ssStoreID, leLegDnStart + 1, (ssStoreID.Length - (leLegDnStart + 2)));
                                    ssArraynum = 1;
                                }
                            }
                            NameResolutionCollection ncCol = service.ResolveName(lnLegDN, ResolveNameSearchLocation.DirectoryOnly, true);
                            if (ncCol.Count > 0)
                            {
                                //Console.WriteLine(ncCol[0].Contact.DisplayName);
                                FolderId SharedCalendarId = new FolderId(WellKnownFolderName.Calendar, ncCol[0].Mailbox.Address);
                                // Folder SharedCalendaFolder = Folder.Bind(service, SharedCalendarId);
                                CalendarFolder calendar = CalendarFolder.Bind(service, new FolderId(WellKnownFolderName.Calendar, ncCol[0].Mailbox.Address), new PropertySet());
                                rtList.Add(ncCol[0].Contact.DisplayName, ncCol[0].Mailbox.Address);


                            }

                        }
                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine(exception.Message);
                    }

                }
            }
            return rtList;
        }

        }
}


