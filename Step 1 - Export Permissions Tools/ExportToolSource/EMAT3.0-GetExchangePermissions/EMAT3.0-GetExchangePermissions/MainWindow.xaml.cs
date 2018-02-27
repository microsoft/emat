using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using System.Threading;
using ActiveDs;
using System.Collections;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Collections.ObjectModel;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using System.Security.AccessControl;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.Principal;

namespace EMAT3._0_GetExchangePermissions { 
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }


    static List<string> Gruppi = new List<string>();

    static string sDomain;
    static string domainFQDN;
    static string username;
    static string password;
    static Object distinguishedName;
    static System.IO.StreamWriter file;
    static System.IO.StreamWriter filemp;
    static System.IO.StreamWriter filead;
    static Hashtable exRighthash = new Hashtable();
    static string WebServicesURL;

    /*  static void GetGroups(PrincipalContext ctx)
      {


          GroupPrincipal qbeGroup = new GroupPrincipal(ctx);
          PrincipalSearcher srch = new PrincipalSearcher(qbeGroup);

          // find all matches
          foreach (var found in srch.FindAll())
          {
              // do whatever here - "found" is of type "Principal" - it could be user, group, computer.....
              //       Console.WriteLine(found.Name);
              // found.StructuralObjectClass
              if (!found.DistinguishedName.ToString().Contains("CN=Builtin"))
              {
                  Gruppi.Add(found.SamAccountName);
              }

          }
      }
      */


    static void GetUsers(PrincipalContext ctx)
    {

        string pathNameDomain = "LDAP://" + sDomain + "/" + distinguishedName.ToString();

        var direcotyEntry = new DirectoryEntry(pathNameDomain, username, password);
        var directorySearcher = new DirectorySearcher(direcotyEntry)
        {
            Filter = "(|(|(|(msExchRecipientTypeDetails=1)(msExchRecipientTypeDetails=4))(msExchRecipientTypeDetails=2))(msExchRecipientTypeDetails=8))"
        };

        directorySearcher.PropertiesToLoad.Add("msExchRecipientTypeDetails");
        directorySearcher.PropertiesToLoad.Add("distinguishedname");
        directorySearcher.PropertiesToLoad.Add("mail");
        directorySearcher.PropertiesToLoad.Add("objectSid");
        directorySearcher.PropertiesToLoad.Add("DisplayName");
        directorySearcher.PropertiesToLoad.Add("SamAccountName");
        directorySearcher.PropertiesToLoad.Add("msExchRecipientDisplayType");
        directorySearcher.SizeLimit = 2000;
        directorySearcher.PageSize = 2000;
        var searchResults = directorySearcher.FindAll();


        foreach (SearchResult searchResult in searchResults)
        {
            var row = new UserInfo();
            row.Distinguishedname = searchResult.Properties["distinguishedname"][0].ToString();

            var temp = searchResult.Properties["mail"];
            if (temp.Count != 0)
            {
                row.Mail = temp[0].ToString();
            }


            var sidBytes = searchResult.Properties["objectSid"][0] as byte[];
            var sid = new SecurityIdentifier(sidBytes, 0).ToString();
            row.ObjectSID = sid.ToString();

            var temp2 = searchResult.Properties["DisplayName"];
            if (temp2.Count != 0)
            {
                row.DisplayName = searchResult.Properties["DisplayName"][0].ToString();
            }

            var temp3 = searchResult.Properties["msExchRecipientTypeDetails"];
            if (temp3.Count != 0)
            {
                row.msExchRecipientTypeDetails = searchResult.Properties["msExchRecipientTypeDetails"][0].ToString();
            }
            row.SamAccountName = searchResult.Properties["samaccountname"][0].ToString();


            var temp4 = searchResult.Properties["msExchRecipientDisplayType"];
            if (temp4.Count != 0)
            {
                row.RecipientType = searchResult.Properties["msExchRecipientDisplayType"][0].ToString();
            }

            Utenti.Add(row);


        }
        Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INIT::found " + Utenti.Count + " Users");
        direcotyEntry.Dispose();
        directorySearcher.Dispose();
        searchResults.Dispose();


        // "Identity","Name","DisplayName","SamAccountName","RecipientType","RecipientTypeDetails","WindowsEmailAddress"

        string stringFilePath = "ExportData\\mailboxes.csv";
        System.IO.TextWriter writer = File.CreateText(stringFilePath);

        string OutputLine = "Identity,Name,DisplayName,SamAccountName,RecipientType,RecipientTypeDetails,WindowsEmailAddress";
        writer.WriteLine(OutputLine);

        foreach (var row in Utenti)
        {
            writer.WriteLine("," + row.DisplayName + "," + row.DisplayName + "," + row.SamAccountName + "," + row.RecipientType + "," + row.RecipientTypeDetails + "," + row.Mail);

        }
        writer.Close();


    }




    static void GetGroupsInfo(PrincipalContext ctx)
    {

        string pathNameDomain = "LDAP://" + sDomain + "/" + distinguishedName.ToString();

        var direcotyEntry = new DirectoryEntry(pathNameDomain, username, password);
        var directorySearcher = new DirectorySearcher(direcotyEntry)
        {
            //      Filter = "(&(objectClass=group)(msExchRecipientDisplayType=1073741833))"
            Filter = "((objectClass=group))"

        };

        directorySearcher.PropertiesToLoad.Add("msExchRecipientTypeDetails");
        directorySearcher.PropertiesToLoad.Add("distinguishedname");
        directorySearcher.PropertiesToLoad.Add("DisplayName");
        directorySearcher.PropertiesToLoad.Add("mail");
        directorySearcher.PropertiesToLoad.Add("objectSid");
        directorySearcher.PropertiesToLoad.Add("mailNickname");
        directorySearcher.PropertiesToLoad.Add("samaccountname");
        directorySearcher.SizeLimit = 2000;
        directorySearcher.PageSize = 2000;
        var searchResults = directorySearcher.FindAll();


        foreach (SearchResult searchResult in searchResults)
        {
            var row = new GroupInfo();
            row.Distinguishedname = searchResult.Properties["distinguishedname"][0].ToString();
            row.samaccountname = searchResult.Properties["samaccountname"][0].ToString();

            var sidBytes = searchResult.Properties["objectSid"][0] as byte[];
            var sid = new SecurityIdentifier(sidBytes, 0).ToString();
            row.ObjectSID = sid.ToString();


            var temp = searchResult.Properties["mail"];
            if (temp.Count != 0)
            {
                row.Mail = temp[0].ToString();
            }


            var temp2 = searchResult.Properties["DisplayName"];
            if (temp2.Count != 0)
            {
                row.DisplayName = searchResult.Properties["DisplayName"][0].ToString();
            }

            var temp3 = searchResult.Properties["msExchRecipientTypeDetails"];
            if (temp3.Count != 0)
            {
                row.msExchRecipientTypeDetails = searchResult.Properties["msExchRecipientTypeDetails"][0].ToString();
            }



            var temp4 = searchResult.Properties["mailNickname"];
            if (temp4.Count != 0)
            {
                row.mailNickname = searchResult.Properties["mailNickname"][0].ToString();
            }

            GruppiInfo.Add(row);

        }
        Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INIT::found " + GruppiInfo.Count + " Groups");


        direcotyEntry.Dispose();
        directorySearcher.Dispose();
        searchResults.Dispose();



        string stringFilePath = "ExportData\\groups.csv";
        System.IO.TextWriter writer = File.CreateText(stringFilePath);

        string OutputLine = "Name,Alias,DisplayName,WindowsEmailAddress,SamAccountName,GroupType";
        writer.WriteLine(OutputLine);

        foreach (var row in GruppiInfo)
        {
            writer.WriteLine(row.DisplayName + "," + row.mailNickname + "," + row.DisplayName + "," + row.Mail + "," + row.samaccountname + "," + row.msExchRecipientTypeDetails);

        }
        writer.Close();


    }



    static void GetFoldersPermissions(string utente)
    {

        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
        service.Credentials = new NetworkCredential(username, password);
        //    service.AutodiscoverUrl(user1Email);
        ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;


        service.Url = new Uri(WebServicesURL);
        // Create a property set to use for folder binding.
        PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.Permissions);

        var mailbox = new Mailbox(utente);


        FolderView view = new FolderView(10);
        view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
        view.PropertySet.Add(FolderSchema.DisplayName);
        SearchFilter searchFilter = new SearchFilter.IsGreaterThan(FolderSchema.TotalCount, 0);
        view.Traversal = FolderTraversal.Deep;


        FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.Root, searchFilter, view);


        foreach (Folder folderSearch in findFolderResults)
        {



            try
            {
                Console.WriteLine(folderSearch.DisplayName);

                GetFolderPermissions(utente, folderSearch.DisplayName.ToLower());
            }
            catch (Exception e) { Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::ERR::FolderPermissions::FAILED::Cannot bind to mailbox " + utente + " --> " + e.Message); }

        }



    }


    public class Book
    {
        public string Identity { get; set; }
        public string User { get; set; }
        public string Folder { get; set; }
        public string Permission { get; set; }
    }

    public static IEnumerable<Book> StreamBooks(string uri)
    {

        string title = null;
        string author = null;

        yield return new Book() { Identity = title, User = author };

    }


    static void GetFolderPermissions(string utente, string folderId)
    {
        NetworkCredential userCredentials = new NetworkCredential(username, password);

        var getFolderSOAPRequest =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n" +
        "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\n" +
        "              xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\"\n" +
        "             xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\"\n" +
        "              xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">\n" +
        " <soap:Header>\n" +
        "   <t:RequestServerVersion Version=\"Exchange2013_SP1\" />\n" +
        " </soap:Header>\n" +
        " <soap:Body>\n" +
        "   <m:GetFolder>\n" +
        "     <m:FolderShape>\n" +
        "       <t:BaseShape>IdOnly</t:BaseShape>\n" +
        "       <t:AdditionalProperties>\n" +
        "         <t:FieldURI FieldURI=\"folder:PermissionSet\" />\n" +
         "       </t:AdditionalProperties>\n" +
         "     </m:FolderShape>\n" +
         "     <m:FolderIds>\n" +
         "       <t:DistinguishedFolderId Id=\"" + folderId + "\">\n" +
         "<t:Mailbox> <t:EmailAddress>" + utente + "</t:EmailAddress> </t:Mailbox>" +
         "</t:DistinguishedFolderId>" +
         "     </m:FolderIds>\n" +
         "   </m:GetFolder>\n" +
         " </soap:Body>\n" +
        "</soap:Envelope>\n";

        // Write the get folder operation request to the console and log file.


        //var getFolderRequest = WebRequest.CreateHttp(Office365WebServicesURL);
        var getFolderRequest = WebRequest.CreateHttp(WebServicesURL);
        ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
        getFolderRequest.AllowAutoRedirect = false;
        getFolderRequest.Credentials = userCredentials;
        getFolderRequest.Method = "POST";
        getFolderRequest.ContentType = "text/xml";


        var requestWriter = new StreamWriter(getFolderRequest.GetRequestStream());
        requestWriter.Write(getFolderSOAPRequest);
        requestWriter.Close();

        try
        {
            var getFolderResponse = (HttpWebResponse)(getFolderRequest.GetResponse());
            if (getFolderResponse.StatusCode == HttpStatusCode.OK)
            {
                var responseStream = getFolderResponse.GetResponseStream();
                XElement responseEnvelope = XElement.Load(responseStream);
                if (responseEnvelope != null)
                {
                    // Write the response to the console and log file.

                    StringBuilder stringBuilder = new StringBuilder();
                    XmlWriterSettings settings = new XmlWriterSettings();
                    settings.Indent = true;
                    XmlWriter writer = XmlWriter.Create(stringBuilder, settings);
                    responseEnvelope.Save(writer);
                    writer.Close();
                    //         Console.WriteLine(responseEnvelope);

                    var Folders = responseEnvelope.Descendants("Folders");
                    //         var Folders = responseEnvelope.Descendants("Folder");
                    //       var Folders = responseEnvelope.Descendants("Folder");




                    foreach (var book in StreamBooks(responseEnvelope.ToString()))
                    {
                        Console.WriteLine("Identity, User: {0}, {1}", book.Identity, book.User);
                    }




                }
            }
        }
        catch (WebException ex)
        {
            Console.WriteLine(ex.Message);
        }
        catch (ApplicationException ex)
        {
            Console.WriteLine(ex.Message);
        }


    }




    private static void GetSendas(UserInfo utente)
    {
        string pathNameDomain = "LDAP://" + sDomain + "/" + utente.Distinguishedname;
        var direcotyEntry = new DirectoryEntry(pathNameDomain, username, password);
        var directorySearcher = new DirectorySearcher(direcotyEntry);
        directorySearcher.PropertiesToLoad.Add("msExchRecipientTypeDetails");
        directorySearcher.PropertiesToLoad.Add("distinguishedname");
        directorySearcher.PropertiesToLoad.Add("mail");
        var res = directorySearcher.FindOne();

        DirectoryEntry ssStoreObj = res.GetDirectoryEntry();
        ActiveDirectorySecurity StoreobjSec = ssStoreObj.ObjectSecurity;
        AuthorizationRuleCollection Storeacls = StoreobjSec.GetAccessRules(true, true, typeof(System.Security.Principal.SecurityIdentifier));
        foreach (ActiveDirectoryAccessRule ace in Storeacls)
        {
            if (ace.IdentityReference.Value != "S-1-5-7" & ace.IdentityReference.Value != "S-1-1-0" & ace.IsInherited != true & ace.IdentityReference.Value != "S-1-5-10")
            {
                if (ace.ActiveDirectoryRights.ToString() == "ExtendedRight")
                {
                    bool found = false;

                    try
                    {
                        filead.WriteLine(utente.Mail + "," + Utenti.Find(x => x.ObjectSID.Contains(ace.IdentityReference.Value)).Mail + ",SendAS," + exRighthash[ace.ObjectType.ToString()].ToString() + ",,");
                        Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::SendAS::OK::SendAS permission of --> " + utente.Mail + " exported successfully");
                        found = true;
                    }
                    catch
                    {

                    }

                    try
                    {
                        filead.WriteLine(utente.Mail + "," + GruppiInfo.Find(x => x.ObjectSID.Contains(ace.IdentityReference.Value)).samaccountname + ",SendAS," + exRighthash[ace.ObjectType.ToString()].ToString() + ",,");
                        Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::SendAS::OK::SendAS permission of --> " + utente.Mail + " exported successfully");
                        found = true;
                    }
                    catch
                    {

                    }

                    if (!found)
                    {
                        Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::WRN::Cannot resolve SID " + ace.IdentityReference.Value);
                    }

                }

            }


        }



    }



    static void GetMBXPermissions(UserInfo utente)
    {


        DirectoryEntry ent = new DirectoryEntry("LDAP://" + sDomain + "/" + utente.Distinguishedname.ToString(), username, password);

        SecurityDescriptor sd = (SecurityDescriptor)ent.Properties["msexchmailboxsecuritydescriptor"].Value;
        AccessControlList acl = (AccessControlList)sd.DiscretionaryAcl;

        foreach (AccessControlEntry ace in (IEnumerable)acl)
        {
            //         Console.WriteLine("Trustee: {0}", ace.Trustee);
            //       Console.WriteLine("AccessMask: {0}", ace.AccessMask);
            //     Console.WriteLine("Access Type: {0}", ace.AceType);
            //   Console.WriteLine("InheritedObjectType: {0}", ace.InheritedObjectType);



            // || ace.InheritedObjectType != null
            if (ace.Trustee != "NT AUTHORITY\\SELF")
            {

                switch (ace.AccessMask)
                {

                    case 131073:
                        bool found = false;
                        try
                        {
                            string find = Utenti.Find(x => x.ObjectSID.Contains(ace.Trustee)).Mail;
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,FullAccess,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ReadPermission,,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {

                        }


                        try
                        {
                            string find = GruppiInfo.Find(x => x.ObjectSID.Contains(ace.Trustee)).samaccountname;
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,FullAccess,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ReadPermission,,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {

                        }

                        if (!found)
                        {
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::WRN::Cannot resolve SID " + ace.Trustee);
                        }


                        break;

                    case 131072:
                        found = false;
                        try
                        {
                            string find = Utenti.Find(x => x.ObjectSID.Contains(ace.Trustee)).Mail;
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,FullAccess,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ReadPermission,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,DeleteItem,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ChangePermission,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ChangeOwner,,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {

                        }


                        try
                        {
                            string find = GruppiInfo.Find(x => x.ObjectSID.Contains(ace.Trustee)).samaccountname;
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,FullAccess,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ReadPermission,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,DeleteItem,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ChangePermission,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ChangeOwner,,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {

                        }

                        if (!found)
                        {
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::WRN::Cannot resolve SID " + ace.Trustee);
                        }


                        break;


                    case 983041:
                        found = false;
                        try
                        {
                            string find = Utenti.Find(x => x.ObjectSID.Contains(ace.Trustee)).Mail;
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,FullAccess,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ReadPermission,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,DeleteItem,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ChangePermission,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ChangeOwner,,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {

                        }


                        try
                        {
                            string find = GruppiInfo.Find(x => x.ObjectSID.Contains(ace.Trustee)).samaccountname;
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,FullAccess,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ReadPermission,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,DeleteItem,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ChangePermission,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,ChangeOwner,,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {

                        }

                        if (!found)
                        {
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::WRN::Cannot resolve SID " + ace.Trustee);
                        }


                        break;



                    case 65537:
                        found = false;
                        try
                        {
                            string find = Utenti.Find(x => x.ObjectSID.Contains(ace.Trustee)).Mail;
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,FullAccess,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,DeleteItem,,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {

                        }


                        try
                        {
                            string find = GruppiInfo.Find(x => x.ObjectSID.Contains(ace.Trustee)).samaccountname;
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,FullAccess,,");
                            filemp.WriteLine(utente.Mail + "," + find + ",MBX,DeleteItem,,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {

                        }

                        if (!found)
                        {
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::WRN::Cannot resolve SID " + ace.Trustee);
                        }


                        break;



                    case 1:
                        found = false;

                        try
                        {
                            filemp.WriteLine(utente.Mail + "," + Utenti.Find(x => x.ObjectSID.Contains(ace.Trustee)).Mail + ",MBX,FullAccess,,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {

                        }


                        try
                        {

                            filemp.WriteLine(utente.Mail + "," + GruppiInfo.Find(x => x.ObjectSID.Contains(ace.Trustee)).samaccountname + ",MBX,FullAccess,,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {

                        }

                        if (!found)
                        {
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::WRN::Cannot resolve SID " + ace.Trustee);
                        }

                        break;


                    default:
                        try
                        {
                            filemp.WriteLine(utente.Mail + "," + Utenti.Find(x => x.ObjectSID.Contains(ace.Trustee)).Mail + ",MBX," + ace.AccessMask + ",,");
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::MBX::OK::MBX permission of --> " + utente.Mail + " exported successfully");
                            found = true;
                        }
                        catch
                        {
                            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::WRN::Cannot resolve SID " + ace.Trustee);
                        }
                        break;
                }

            }
        }

    }
    // 1 FullAccess
    // 131073   FullAccess + ReadPermission
    // 131072   FullAccess, DeleteItem, ReadPermission, ChangePermission, ChangeOwner
    // 983041   FullAccess, DeleteItem, ReadPermission, ChangePermission, ChangeOwner
    // 65537    FullAccess, DeleteItem
    // InheritedObjectType controllare se è disponibile in modo da filtrare (filtro inserito da verificare)



    static void GetAllUsersFromGroup(string group, string domain)
    {
        List<string> retVal = new List<string>();
        DirectoryEntry entry = new DirectoryEntry("LDAP://" + sDomain + "/" + distinguishedName.ToString(), username, password);
        DirectorySearcher searcher = new DirectorySearcher("(&(objectCategory=group)(samaccountname=" + group + "))");
        searcher.SearchRoot = entry;
        searcher.SearchScope = SearchScope.Subtree;
        SearchResult result = searcher.FindOne();

        if (result != null)
        {
            foreach (string member in result.Properties["member"])
            {
                string temp = String.Concat("LDAP://", sDomain, "/", member.ToString());
                DirectoryEntry de = new DirectoryEntry(temp, username, password);
                if (de.Properties["objectClass"].Contains("group") || de.Properties["objectClass"].Contains("user") && de.Properties["cn"].Count > 0)
                {
                    Console.Write(group + "----->");
                    Console.WriteLine(de.Properties["samaccountname"][0].ToString());
                    file.WriteLine(de.Properties["samaccountname"][0].ToString() + "," + group + "," + domain);
                }

            }

        }

    }



    static string GetNetbiosDomainName(string dnsDomainName)
    {

        string temp = "LDAP://" + sDomain + "/CN=Partitions,CN=Configuration," + distinguishedName.ToString();
        DirectoryEntry entry = new DirectoryEntry(temp, username, password);
        DirectorySearcher searcher = new DirectorySearcher("(&(objectcategory=Crossref)(dnsRoot=" + dnsDomainName + "))");
        searcher.SearchRoot = entry;
        searcher.SearchScope = SearchScope.Subtree;
        searcher.Filter = "nETBIOSName=*";
        SearchResult result = searcher.FindOne();

        if (result != null)
            return result.Properties["nETBIOSName"][0].ToString();
        else return null;


    }
    class UserInfo
    {
        public string Distinguishedname { get; set; }
        public string Mail { get; set; }
        public string ObjectSID { get; set; }
        public string msExchRecipientTypeDetails { get; set; }

        public string DisplayName { get; set; }
        public string SamAccountName { get; set; }

        public string RecipientType { get; set; }
        public string RecipientTypeDetails { get; set; }
    }


    class GroupInfo
    {
        public string Distinguishedname { get; set; }
        public string Mail { get; set; }
        public string ObjectSID { get; set; }
        public string msExchRecipientTypeDetails { get; set; }
        public string samaccountname { get; set; }
        public string DisplayName { get; set; }
        public string mailNickname { get; set; }
    }

    static List<UserInfo> Utenti = new List<UserInfo>();
    static List<GroupInfo> GruppiInfo = new List<GroupInfo>();


    private async void button_Click(object sender, RoutedEventArgs e)
    {


        string subPath = "ExportData";

        bool exists = System.IO.Directory.Exists(subPath);

        if (!exists)
            System.IO.Directory.CreateDirectory(subPath);

        Trace.Listeners.Add(new TextWriterTraceListener("trace.log"));
        Trace.AutoFlush = true;


        Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::Starting");
        ContextOptions options = ContextOptions.SimpleBind;
        textBlock_log.Text = "Fill the fields above and press the start button";
        button.IsEnabled = false;
        textBox_DCIP.IsEnabled = false;
        textBox_domainFQDN.IsEnabled = false;
        textBox_username.IsEnabled = false;
        passwordBox.IsEnabled = false;
        textBox_EWS.IsEnabled = false;
        GetGroupMembership.IsEnabled = false;
        CheckboxGetFolderPermissions.IsEnabled = false;
        GetMailboxPermissions.IsEnabled = false;
        CheckboxGetSendas.IsEnabled = false;



        sDomain = textBox_DCIP.Text;
        domainFQDN = textBox_domainFQDN.Text;
        username = textBox_username.Text;
        //    username = "EMEA\\administrator";

        WebServicesURL = "https://" + textBox_EWS.Text + "/EWS/Exchange.asmx";


        password = passwordBox.Password;
        //  password = "Password.123";

        textBlock_log.Text = "Getting Domain NETBIOS";
        Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::Getting Domain NETBIOS");

        try
        {
            DirectoryEntry RootDirEntry = new DirectoryEntry("LDAP://" + sDomain + "/RootDSE", username, password);
            distinguishedName = RootDirEntry.Properties["defaultNamingContext"].Value;

        }
        catch (Exception k)
        {
            Trace.WriteLine(k);
            textBlock_log.Text = "Cannot RootDSE....check Domain FQDN and Domain Controller IP";
        }

        string ret;
        try
        {
            ret = GetNetbiosDomainName(domainFQDN);
        }
        catch (Exception k)
        {
            Trace.WriteLine(k);
            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::WRN::Cannot find Domain NETBIOS....check Domain FQDN and Domain Controller IP....Using DUMMY");
            textBlock_log.Text = "Cannot find Domain NETBIOS....check Domain FQDN and Domain Controller IP....Using DUMMY";
            ret = "DUMMY";
        }



        /*
                    PrincipalContext ctx = new PrincipalContext(ContextType.Domain, sDomain, null, options, username, password);
                    await System.Threading.Tasks.Task.Run(() => GetGroupInfo(ctx));
                    ctx.Dispose();
                    */

        int i = 0;


        PrincipalContext ctx = new PrincipalContext(ContextType.Domain, sDomain, null, options, username, password);
        textBlock_log.Text = "Reading Groups from Active Directory...it may take a while";
        Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::Reading Groups from Active Directory");
        //        await System.Threading.Tasks.Task.Run(() => GetGroupsInfo(ctx));
        Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::Reading Groups from Active Directory...DONE");
        textBlock_log.Text = "Reading Groups from Active Directory.....DONE";

        if (GetGroupMembership.IsChecked.Value)
        {

            file = new System.IO.StreamWriter("ExportData\\groupsmembership.csv");
            file.AutoFlush = true;
            file.WriteLine("Samaccountname,Group,Domain");


            textBlock_log.Text = "Starting reading membership";
            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::Starting reading membership");

            int totalegruppi = GruppiInfo.Count();

            textBlock_log.Text = "Found " + Gruppi.Count() + " groups";
            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::Found " + Gruppi.Count() + " groups");

            foreach (var gruppo in GruppiInfo)
            {
                i++;
                int percent = 100 * i / totalegruppi;
                try
                {
                    Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::Reading membership of --> " + gruppo.samaccountname + " -- Total progress: " + percent + "%");
                    await System.Threading.Tasks.Task.Run(() => GetAllUsersFromGroup(gruppo.samaccountname, ret));

                }
                catch (Exception ex)
                {
                    Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::FAILED::I was reading ---> " + gruppo);
                    Trace.WriteLine(ex);
                }
                textBlock_log.Text = "Working on: Reading Group Membership: " + percent + "% completed , " + i + " of " + totalegruppi;
            }

            ctx.Dispose();
        }




        if (GetMailboxPermissions.IsChecked.Value || CheckboxGetSendas.IsChecked.Value || CheckboxGetFolderPermissions.IsChecked.Value)
        {
            DirectoryEntry rootdse = new DirectoryEntry("LDAP://" + sDomain + "/RootDSE", username, password);
            DirectoryEntry cfg = new DirectoryEntry("LDAP://" + sDomain + "/" + rootdse.Properties["configurationnamingcontext"].Value, username, password);
            DirectoryEntry exRights = new DirectoryEntry("LDAP://" + sDomain + "/cn =Extended-rights," + rootdse.Properties["configurationnamingcontext"].Value, username, password);

            foreach (DirectoryEntry chent in exRights.Children)
            {
                if (exRighthash.ContainsKey(chent.Properties["rightsGuid"].Value) == false)
                {
                    exRighthash.Add(chent.Properties["rightsGuid"].Value, chent.Properties["DisplayName"].Value);
                }
            }
            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::INIT::READ::Got Extended Rights definition");

            ctx = new PrincipalContext(ContextType.Domain, sDomain, null, options, username, password);
            textBlock_log.Text = "Reading users list......it may take a while!";
            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::Reading users list......it may take a while!");
            await System.Threading.Tasks.Task.Run(() => GetUsers(ctx));

        }

        int totaleutenti = Utenti.Count();



        filemp = new System.IO.StreamWriter("ExportData\\mp.csv");
        filemp.Flush();
        filemp.AutoFlush = true;
        filemp.WriteLine("Identity,User,PermissionType,AccessRight,ExtendedRight,FolderPath");



        filead = new System.IO.StreamWriter("ExportData\\ad.csv");
        filead.Flush();
        filead.AutoFlush = true;
        filead.WriteLine("Identity,User,PermissionType,AccessRight,ExtendedRight,FolderPath");
        ParallelOptions opts = new ParallelOptions { MaxDegreeOfParallelism = Convert.ToInt32(Math.Ceiling((Environment.ProcessorCount * 0.70) * 1.0)) };


        i = 0;

        bool fp = CheckboxGetFolderPermissions.IsChecked.Value;
        bool sendas = CheckboxGetSendas.IsChecked.Value;
        bool mp = GetMailboxPermissions.IsChecked.Value;

        //   Parallel.ForEach(Utenti, opts, (utente) =>
        foreach (UserInfo utente in Utenti)
        {

            i++;
            int percent = 100 * i / totaleutenti;
            try
            {
                System.IO.File.WriteAllText("resume.chk", utente.Mail);
            }
            catch
            {
                Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::WRN::Resume check file is locked");
            }

            Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::READ::Reading permission of --> " + utente.Mail + " -- Total progress: " + percent + "%, " + i + " of " + totaleutenti);
            textBlock_log.Text = "Working on: Reading permissions: " + percent + "% completed, " + i + " of " + totaleutenti;

            if (fp)
            {
                await System.Threading.Tasks.Task.Run(() => GetFoldersPermissions(utente.Mail));
            }

            if (sendas)
            {
                await System.Threading.Tasks.Task.Run(() => GetSendas(utente));
            }


            if (mp)
            {
                await System.Threading.Tasks.Task.Run(() => GetMBXPermissions(utente));
            }


            /*     this.Dispatcher.Invoke(() =>       {                });*/
        }


        try
        {
            filemp.Close();
            filead.Close();
        }
        catch { }

        Trace.WriteLine(DateTime.Now.ToString("yyyyMMddHHmmss") + "::INF::Job Finished");
        textBlock_log.Text = "Job Finished";

        button.IsEnabled = true;
        textBox_DCIP.IsEnabled = true;
        textBox_domainFQDN.IsEnabled = true;
        textBox_username.IsEnabled = true;
        passwordBox.IsEnabled = true;

        GetGroupMembership.IsEnabled = true;
        CheckboxGetFolderPermissions.IsEnabled = false;
        GetMailboxPermissions.IsEnabled = true;
        CheckboxGetSendas.IsEnabled = true;
        textBox_EWS.IsEnabled = true;

        //Trace.Unindent();
    }

    private void textBox_username_GotFocus(object sender, RoutedEventArgs e)
    {
        textBox_username.Clear();
        textBox_username.GotFocus -= textBox_username_GotFocus;
        passwordBox.IsEnabled = true;
    }

    private void passwordBox_GotFocus(object sender, RoutedEventArgs e)
    {
        passwordBox.Clear();
        passwordBox.GotFocus -= passwordBox_GotFocus;
        textBox_domainFQDN.IsEnabled = true;

    }

    private void textBox_domainFQDN_GotFocus(object sender, RoutedEventArgs e)
    {
        textBox_domainFQDN.Clear();
        textBox_domainFQDN.GotFocus -= textBox_domainFQDN_GotFocus;

        textBox_DCIP.IsEnabled = true;

    }

    private void textBox_DCIP_GotFocus(object sender, RoutedEventArgs e)
    {

        textBox_DCIP.Clear();
        try
        {
            IPHostEntry temp = Dns.GetHostEntry(textBox_domainFQDN.Text);
            textBox_DCIP.Text = temp.AddressList[0].ToString();
        }
        catch { textBox_DCIP.Text = "DNS ERROR: DOMAIN NOT FOUND"; }

        textBox_DCIP.GotFocus -= textBox_DCIP_GotFocus;
        textBlock_log.Text = "Ready to start";
    }

}
}