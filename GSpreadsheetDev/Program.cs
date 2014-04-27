using Google.GData.Client;
using Google.GData.Spreadsheets;
using System;
using System.Collections.Generic;
using GSpreadsheetDev.LarimerParcelInfo;
using System.Configuration;


namespace GSpreadsheetDev
{
    class Program
    {
        //http://www.larimer.org/databases/api.htm
        //https://developers.google.com/google-apps/spreadsheets/#working_with_cell-based_feeds
        //https://developers.google.com/google-apps/spreadsheets/#working_with_cell-based_feeds
        //http://www.larimer.org/webservices/PropertyInformation.cfc?wsdl

        public static string userName = ConfigurationSettings.AppSettings["gAccount"];
        public static string password = ConfigurationSettings.AppSettings["gPassword"];
        public static string sheetService = ConfigurationSettings.AppSettings["sheetService"];
        public static string sheetTitle = ConfigurationSettings.AppSettings["sheetTitle"];
        public static int worksheetNumber = Int32.Parse(ConfigurationSettings.AppSettings["worksheetNumber"].ToString());
        public static int parcelColumnNumber = Int32.Parse(ConfigurationSettings.AppSettings["parcelColumnNumber"].ToString());

        static void Main()
        {
            
            // spreadsheet service and auth
            SpreadsheetsService _service = new SpreadsheetsService(sheetService);
            _service.setUserCredentials(userName, password);

            // new spreadsheet query
            SpreadsheetQuery query = new SpreadsheetQuery();
            query.Title = sheetTitle; // will find the spreadsheet using title

            // make the request
            SpreadsheetFeed feed = _service.Query(query);

            // get spreadsheet 
            SpreadsheetEntry spreadsheet = (SpreadsheetEntry)feed.Entries[0];

            // request the spreadsheets worksheet feeds
            WorksheetFeed wsFeeds = spreadsheet.Worksheets;

            // target worksheet
            WorksheetEntry worksheet = (WorksheetEntry)wsFeeds.Entries[worksheetNumber];

            // need request url for the list feed
            AtomLink listFeedLink = worksheet.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);

            // fetch a list
            ListQuery listQuery = new ListQuery(listFeedLink.HRef.ToString());
            ListFeed parcelList = _service.Query(listQuery);


            // fetch list of members
            WorksheetEntry memberWorksheet = (WorksheetEntry)wsFeeds.Entries[0];
            AtomLink memberListFeedLink = memberWorksheet.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);
            ListQuery memberListQuery = new ListQuery(memberListFeedLink.HRef.ToString());
            ListFeed memberList = _service.Query(memberListQuery);

            // loop through each entry and update the legal cell for each row
            int i = 0;
            foreach (ListEntry entry in parcelList.Entries)
            {
                i++;
                PropertyInfo property = new PropertyInformation().GetPropertyInfo("", entry.Elements[parcelColumnNumber].Value.ToString());
                Parcel parcel = new Parcel(property, memberList);

                Console.WriteLine("{0} => {1}, {2}", i,  parcel.Number, parcel.IsMember);

                // look at each element in entry
                foreach (ListEntry.Custom element in entry.Elements)
                {
                    
                    if (element.LocalName == "legal")
                    {
                        element.Value = parcel.Legal;
                    }

                    if (element.LocalName == "subdivision")
                    {
                        element.Value = parcel.Subdivision;
                    }

                    if (element.LocalName == "filing")
                    {
                        element.Value = parcel.Filing;
                    }

                    if (element.LocalName == "ismember")
                    {
                        element.Value = parcel.IsMember == true ? "1" : "0";
                    }

                    if (element.LocalName == "grantsigned")
                    {
                        element.Value = parcel.GrantSigned;
                    }

                    if (element.LocalName == "grantfiled")
                    {
                        element.Value = parcel.GrantFiled;
                    }

                    if (element.LocalName == "memberid")
                    {
                        element.Value = parcel.MemberId;
                    }

                    if (element.LocalName == "annexmarkercolor")
                    {
                        var color = "";
                        if (parcel.IsMember)
                        {
                            color = "small_blue";
                        }
                        else
                        {
                            color = parcel.IsAnnexed ? "small_red" : "small_yellow";
                        }
                         
                         element.Value = color;
                    }


                    // mailing address
                    if (element.LocalName == "mailaddress1")
                    {
                        element.Value = parcel.MailAddress1;
                    }

                    if (element.LocalName == "mailaddress2")
                    {
                        element.Value = parcel.MailAddress2;
                    }

                    if (element.LocalName == "mailaddresscity")
                    {
                        element.Value = parcel.MailAddressCity;
                    }

                    if (element.LocalName == "mailaddressstate")
                    {
                        element.Value = parcel.MailAddressState;
                    }

                    if (element.LocalName == "mailaddresszip")
                    {
                        element.Value = parcel.MailAddressZip;
                    }

                    // 1st owner
                    if (element.LocalName == "owner1")
                    {
                        element.Value = parcel.Owner1;
                    }

                    // 2nd owner
                    if (element.LocalName == "owner2")
                    {
                        element.Value = parcel.Owner2;
                    }
                }

                entry.Update();
            }

            Console.WriteLine("Done, press any key");
            Console.ReadKey();
        }
    }
}
