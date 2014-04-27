using Google.GData.Spreadsheets;
using GSpreadsheetDev.LarimerParcelInfo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
namespace GSpreadsheetDev
{
    public class Parcel
    {
        public string Number { get; set; }
        public string Owner { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Zip { get; set; }
        public string Legal { get; set; }
        public string Subdivision { get; set; }
        public string Filing { get; set; }
        public string GrantSigned { get; set; }
        public string GrantFiled { get; set; }
        public string MemberId { get; set; }
        public bool IsMember { get; set; }
        public bool IsAnnexed { get; set; }
        public string MarkerColor { get; set; }

        // mailing address
        public string MailAddress1 { get; set; }
        public string MailAddress2 { get; set; }
        public string MailAddressCity { get; set; }
        public string MailAddressState { get; set; }
        public string MailAddressZip { get; set; }
        
        // owners according to Larimer
        public string Owner1 { get; set; }
        public string Owner2 { get; set; }
        

        public Parcel() { }

        public Parcel(PropertyInfo p, ListFeed memberList)
        {
            this.Number = p.parcelNumber;
            this.Owner = p.ownerName1;
            this.Address = p.address;
            this.City = p.city;
            this.Zip = p.zipCode;
            this.Legal = p.legal;
            this.Subdivision = !string.IsNullOrWhiteSpace(p.subdivDescr) ? p.subdivDescr : GetSubdivision();
            this.Filing = GetFiling();
            this.IsAnnexed = GetAnnex();
            Membership m = GetLienInfo(memberList);
            this.IsMember = !string.IsNullOrWhiteSpace(m.MemberId) ? true : false;
            this.GrantFiled = !string.IsNullOrWhiteSpace(m.GrantFiled) ? m.GrantSigned : "";
            this.GrantSigned = !string.IsNullOrWhiteSpace(m.GrantSigned) ? m.GrantSigned : "";
            this.MemberId = !string.IsNullOrWhiteSpace(m.MemberId) ? m.MemberId : "";

            // mailing addres
            this.MailAddress1 = !string.IsNullOrWhiteSpace(p.mailAddress1) ? p.mailAddress1 : "";
            this.MailAddress2 = !string.IsNullOrWhiteSpace(p.mailAddress2) ? p.mailAddress2 : "";
            this.MailAddressCity = !string.IsNullOrWhiteSpace(p.mailCity) ? p.mailCity : "";
            this.MailAddressState = !string.IsNullOrWhiteSpace(p.mailState) ? p.mailState : "";
            this.MailAddressZip = !string.IsNullOrWhiteSpace(p.mailZipCode) ? FormatZipCode(p.mailZipCode) : "";

            // owner
            this.Owner1 = this.Owner;
            this.Owner2 = !string.IsNullOrWhiteSpace(p.ownerName2) ? p.ownerName2 : "";


        }

        private string FormatZipCode(string zipCode)
        {
            var regExp = new Regex(@"^\d{5}(-\d{4})?$");

            if (!regExp.IsMatch(zipCode))
            {
                // zip does not match xxxxx or xxxxx-xxxx
                // so we have to format it.. cheaply
                zipCode = zipCode.Substring(0, 5) + "-" + zipCode.Substring(5, 4);
            }
            
            return zipCode;
        }

        private string GetFiling()
        {
            if (string.IsNullOrWhiteSpace(this.Legal))
            {
                return "";
            }

            foreach (var item in Filings())
            {
                if (Legal.ToLower().Contains(item.ToLower()))
                {
                    return item.ToString();
                }
            }
            return "";
        }

        private bool GetAnnex()
        {
            // Foothills Village First Filing, 
            // Foothills Green First Filing, 
            // Village West Sixth Filing, 
            // Village West Seventh Filing, 
            // Village West Eighth Filing, 
            // Village West Ninth Filing, 
            // Aspen Knolls First Filing, 
            // Hill Pond on Spring Creek Second Filing
            // Village West Fifth Filing which are located on 
            // - Independence Road,  
            // - Independence Court, 
            // - and the east side of Constitution Avenue between Independence Road and Winfield Drive 

            if (this.Subdivision.ToLower().Contains("foothills village"))
            {
                return this.Legal.ToLower().Contains("1st");
            }

            if (this.Subdivision.ToLower().Contains("foothills green"))
            {
                return this.Legal.ToLower().Contains("1st");
            }

            if (this.Subdivision.ToLower().Contains("village west"))
            {
                if (this.Legal.ToLower().Contains("5th"))
                {
                    if (this.Address.ToLower().Contains("independence rd"))
                    {
                        return true;
                    }

                    if (this.Address.ToLower().Contains("independence ct"))
                    {
                        return true;
                    }

                    if (this.Address.ToLower().Contains("constitution ave"))
                    {
                        string[] d = this.Address.Split(' ');
                        int n = Int32.Parse(d[0]);
                        return n % 2 == 0;
                    }
                }


                if (this.Legal.ToLower().Contains("6th"))
                {
                    return true;
                }

                if (this.Legal.ToLower().Contains("7th"))
                {
                    return true;
                }

                if (this.Legal.ToLower().Contains("8th"))
                {
                    return true;
                }

                if (this.Legal.ToLower().Contains("9th"))
                {
                    return true;
                }
            }

            if (this.Legal.ToLower().Contains("aspen knolls"))
            {
                return true;
            }

            if (this.Legal.ToLower().Contains("hillpond"))
            {
                return true;
            }
            return false;
        }

        private string[] Filings()
        {
            return new string[] { "1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th", "11th", "12th", "13th", "14th", "15th" };
        }

        private string[] Subdivisions()
        {
            return new string[] { "Foothills Green", "Foothills Village", "Foothills West", "Hill Pond", "Peck Minor", "Aspen Knolls" };
        }

        private string GetSubdivision()
        {
            if (this.Legal.ToLower().Contains("foothills green"))
            {
                return "FOOTHILLS GREEN";
            }

            if (this.Legal.ToLower().Contains("foothills village"))
            {
                return "FOOTHILLS VILLAGE";
            }

            if (this.Legal.ToLower().Contains("foothills west"))
            {
                return "FOOTHILLS WEST";
            }

            if (this.Legal.ToLower().Contains("hllpond"))
            {
                return "HILL POND";
            }

            if (this.Legal.ToLower().Contains("aspen"))
            {
                return "ASPEN KNOLLS";
            }
            return "";
        }

        private Membership GetLienInfo(ListFeed memberList)
        {
            Membership m = new Membership();
            foreach (ListEntry entry in memberList.Entries)
            {
                if (entry.Elements[8].Value.ToString().Replace("-", "") == this.Number)
                {
                    m.GrantFiled = entry.Elements[6].Value.ToString();
                    m.GrantSigned = entry.Elements[5].Value.ToString();
                    m.MemberId = entry.Elements[0].Value.ToString();
                }
            }
            return m;
        }
    }
}
