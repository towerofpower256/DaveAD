using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;

namespace DavoAD
{
    public class ADTools
    {

        /// <summary>
        /// Retrieve a property / attribute's value as a string from a Directory Entry
        /// </summary>
        /// <param name="ent">Directory Entry object to read from</param>
        /// <param name="propName">The property / attribute's name</param>
        /// <returns>Property as a string</returns>
        public static string GetPropString(DirectoryEntry ent, String propName)
        {
            //Retrieves a property as a string
            //REMEMBER: if it returns as an array, use GetPropArray!!!
            return (ent.Properties[propName].Value ?? "").ToString();
        }

        /// <summary>
        /// Retrieve a propert / attribute as an array. Useful for multi-line or multi-value properties
        /// </summary>
        /// <param name="ent">DirEnt object to read from</param>
        /// <param name="propName">Property / attribute name to read</param>
        /// <returns>String array of property</returns>
        public static string[] GetPropStringArray(DirectoryEntry ent, String propName)
        {
            int count = ent.Properties[propName].Count;

            if (count > 0)
            {
                string[] arr = new string[count];
                for (int i = 0; i < count; i++)
                {
                    arr[i] = (ent.Properties[propName][i] ?? "").ToString();
                }

                return arr;
            }
            else
            {
                return (new string[0]);
            }
        }

        /// <summary>
        /// Function to search AD, with most options available
        /// </summary>
        /// <param name="filter">LDAP filter to use</param>
        /// <param name="maxResults">Number of results to return. Use 0 with a value in PageSize to return all results</param>
        /// <param name="pageSize">If maxResults is 0 and a PageSize given, will return all results. Recommended = 500</param>
        /// <param name="ordered">Should the results be ordered?</param>
        /// <param name="orderby">Attribute / Property to order by</param>
        /// <param name="orderDirection">Ascending or Descending ordering</param>
        /// <param name="ldapRoot">DirEnt LDAP connection to search using</param>
        /// <returns>Returns an array of DirEnts found in the search</returns>
        public static DirectoryEntry[] SearchAD(
            string filter,
            DirectoryEntry ldapRoot,
            int maxResults = 0,
            int pageSize = 0,
            bool ordered = false,
            string orderby = "name",
            SortDirection orderDirection = SortDirection.Ascending
            )
        {
            //Prep the searcher
            DirectorySearcher searcher = new DirectorySearcher(ldapRoot);

            //Set some stuff
            searcher.Filter = filter;
            searcher.SizeLimit = maxResults;
            searcher.PageSize = pageSize;

            if (ordered == true)
            {
                searcher.Sort = new SortOption(orderby, orderDirection);
            }

            //Search!
            SearchResultCollection rs = searcher.FindAll();

            DirectoryEntry[] returnArray = new DirectoryEntry[rs.Count];
            for (int i = 0; i < rs.Count; i++)
            {
                returnArray[i] = rs[i].GetDirectoryEntry();
            }
            return returnArray;
        }

        /// <summary>
        /// In case the string isn't prefixed by LDAP:// or LDAPS://, this will add the prefix
        /// If the incorrect prefix is already there, it will replace it
        /// </summary>
        /// <param name="domain">String to input</param>
        /// <param name="ldaps">Use LDAPS? If not, just use LDAP</param>
        /// <returns>Input string, prefixed</returns>
        public static string PrefixDomain(string domain, bool ldaps = false)
        {
            string prefix = (ldaps ? "LDAP://" : "LDAPS://");

            if (domain.Length < 8)
            {//Short domain, probably doesn't have prefix
                return prefix + domain;
            }
            else
            {
                string currentPrefix = domain.Substring(0, 8).ToUpper();
                if (currentPrefix.Substring(0,7) == "LDAP://")
                {
                    //If LDAP prefix detected
                    if (ldaps)
                    {//Should be secure, strip Prefix and return with LDAPS prefix
                        return prefix + domain.Replace("LDAP://", "");
                    }
                    else
                    {
                        return domain; //Already correctly prefixed, nothing needs updating
                    }
                }
                else if (currentPrefix == "LDAPS://")
                {
                    //If LDAPS prefix detected
                    if (ldaps)
                    {
                        return domain; //Already prefixed, nothing needs updating
                    }
                    else
                    {//Should be non-secure, strip LDAPS and replace with propper prefix
                        return prefix + domain.Replace("LDAPS://", "");
                    }
                }
                else
                {//Not prefixed, add propper prefix
                    return prefix + domain;
                }
            }
        }

        /// <summary>
        /// Strips away the domain prefix from a string
        /// </summary>
        /// <param name="domain">Input string</param>
        /// <returns>String with LDAPS:// and LDAPS:// removed</returns>
        public static string PrefixDomainStrip(string domain)
        {
            return domain.Replace("LDAP://", "").Replace("LDAPS://", "");
        }

        /// <summary>
        /// For a long FQDN, gets the first item
        /// </summary>
        /// <param name="x">Input FQDN string</param>
        /// <returns></returns>
        public static string GetFirstItemOfDN(string x)
        {
            return PrefixDomainStrip(x).Split((char)44)[0].Substring(3); //Get the first component and drop the first 3 characters E.g. cn=
        }

        /// <summary>
        /// Turns a GUID object into a string which can be used in LDAP queries
        /// </summary>
        /// <param name="input">GUID object to be encoded</param>
        /// <returns>A GUID string usable in LDAP queries</returns>
        public static string EncodeGUIDForQuery(Guid input)
        {
            string retString = "";
            //string guidString = input.ToString().Replace("-", "");
            byte[] guidBytes = input.ToByteArray();
            for (int i = 0; i < guidBytes.Length; i++)
            {
                //retString += (i % 2 == 0 ? "\\" : "") + guidBytes[i].ToString("X");
                retString += "\\" + guidBytes[i].ToString("X");
            }

            return retString;
            //e.g. 16ae78c7-031a-48f9-a255-8183d7c84ee7
            // becomes \C7\78\AE\16\1A\3\F9\48\A2\55\81\83\D7\C8\4E\E7
            //Don't ask me how it gets moved around like that, it's something about the groupings being reversed. Google it if you're interested.
        }

        /// <summary>
        /// Nice and fleshy AD search, similar to the search in AD Users and Computers
        /// Searches AD by name, firstname, lastname, email and username
        /// Perfect for "Search AD" fields where the fields to search aren't specified
        /// </summary>
        /// <param name="searchString">Search paramater to search for</param>
        /// <returns></returns>
        public static DirectoryEntry[] ComprehensiveADSearch(string searchString, DirectoryEntry LdapObject, int PageSize = 50)
        {
            //Just in case they're retarded, trim
            searchString = searchString.Trim();

            DirectorySearcher searcher = new DirectorySearcher(LdapObject);

            string[] searchStringSplit = searchString.Split((char)32); //Split by space
            if (searchStringSplit.Length > 1) //Two or more components
            {
                searcher.Filter = string.Format("(&(sAMAccountType=805306368)(|(&(givenname={0}*)(sn={1}*))(name={2}*)(mail={2}*)(sAMAccountName={2}*)))",
                    searchStringSplit[0],
                    string.Join(" ", searchStringSplit.Skip(1)),
                    searchString
                    );
            }
            else
            {
                searcher.Filter = string.Format("(&(sAMAccountType=805306368)(|(name={0}*)(mail={0}*)(sAMAccountName={0}*)))", searchString);
            }

            searcher.PageSize = PageSize;

            searcher.PropertiesToLoad.Clear();
            searcher.PropertiesToLoad.Add("cn"); //Just load the CN

            SearchResultCollection results = searcher.FindAll();
            List<DirectoryEntry> returnArray = new List<DirectoryEntry>();
            foreach (SearchResult r in results)
            {
                returnArray.Add(r.GetDirectoryEntry());
            }

            return returnArray.ToArray();
        }

        /// <summary>
        /// Quick AD search to retrieve an object by GUID
        /// </summary>
        /// <param name="searchFor">GUID value to search for</param>
        /// <param name="LdapConnection">DirEnt LDAP connection to search using</param>
        /// <returns>Single DirectoryEntry object or null</returns>
        public static DirectoryEntry GetObjectByGuid(Guid searchFor, DirectoryEntry LdapConnection)
        {
            DirectorySearcher searcher = new DirectorySearcher(LdapConnection);

            searcher.Filter = string.Format("(objectGuid={0})", EncodeGUIDForQuery(searchFor));
            searcher.SizeLimit = 1;

            SearchResult result = searcher.FindOne();

            if (result == null)
            {
                return null;
            }
            else
            {
                return result.GetDirectoryEntry();
            }
        }

        /// <summary>
        /// Quick AD search to retrieve a DirEnt object by username / sAMAccount Name
        /// </summary>
        /// <param name="searchFor">sAMAccount Name to search for</param>
        /// <param name="LdapConnection">DirEnt connection to use</param>
        /// <returns>DirEnt of object or Null</returns>
        public static DirectoryEntry GetObjectBySamaccountName(string searchFor, DirectoryEntry LdapConnection)
        {
            DirectorySearcher searcher = new DirectorySearcher(LdapConnection);

            searcher.Filter = string.Format("(sAMAccountName={0})", searchFor);
            searcher.SizeLimit = 1;

            SearchResult result = searcher.FindOne();

            if (result == null)
            {
                return null;
            }
            else
            {
                return result.GetDirectoryEntry();
            }
        }


        /// <summary>
        /// Sets the password for a DirEnt. Will likely need an LDAP connection with admin credentials given
        /// </summary>
        /// <param name="ent">DirEnt to update password on</param>
        /// <param name="newpass">New password</param>
        public static void SetPassword(DirectoryEntry ent, string newpass)
        {
            //Set pass
            ent.Invoke("SetPassword", new object[] { newpass });

            //Dont have to change password at next login
            const int ADS_UF_PASSWORD_EXPIRED = 0x800000;
            int curval = (int)ent.Properties["userAccountControl"].Value;
            ent.Properties["userAccountControl"].Value = curval & (~ADS_UF_PASSWORD_EXPIRED);

            //Save changes
            ent.CommitChanges();
        }


        /// <summary>
        /// Attempts to auth against AD. If it's successful, nothing will happen. 
        /// If it fails, it will raise an exception with a reason (most likely bad username or password
        /// </summary>
        /// <param name="domain">LDAP string (e.g. LDAPS://my.domain.com or LDAP://DC=my,DC=domain,DC=com</param>
        /// <param name="username">Username</param>
        /// <param name="password">Password</param>
        public static void AuthenticateAgainstAD(string domain, string username, string password)
        {
            var ldap = new DirectoryEntry(domain, username, password);
            var failIfBroken = ldap.NativeObject;
            return; //If the above didn't raise an exception, then authenticate succeeded.
        }

        /// <summary>
        /// Checks if a object (by DN) is a member of group(s). Good for security checks where only members of certain groups are allowed in.
        /// </summary>
        /// <param name="ldap">LDAP connection object</param>
        /// <param name="TargetDN">DN of the object to check for</param>
        /// <param name="GroupDN">String array of group names (CN) to check against</param>
        /// <param name="recurse">Should recurse through sub-groups or just check the target group?</param>
        /// <returns>Was found in one or more of those groups</returns>
        public static bool IsMemberOfGroup(DirectoryEntry ldap, string TargetDN, string[] Groups, bool recurse = true)
        {
            bool found = false;
            foreach (string g in Groups)
            {
                //Resolve the name to DN
                string gName = "";
                DirectoryEntry groupNameResult = GetObjectBySamaccountName(g, ldap);
                if (groupNameResult == null)
                {
                    continue; //Can't resolve this group, ignore it
                }
                else
                {
                    gName = DavoAD.ADTools.GetPropString(groupNameResult, "distinguishedName");
                }

                string queryString = string.Format("(&(distinguishedName={0})({1}={2}))", TargetDN,
                    (recurse ? "memberOf:1.2.840.113556.1.4.1941:" : "memberOf"), //If recurse = True, use fancy MemberOf to recurse through sub-groups. Otherwise, just test immediate group, nothing more
                    gName);
                DirectoryEntry[] results = SearchAD(queryString, ldap, 1);
                if (results.Length > 0)
                {
                    found = true; //True if there's more than 0 results found, false
                    break;
                } 
            }

            return found;
            
        }

        /// <summary>
        /// Class to make AD Timestamps easier
        /// ADTimestamp lolwut = new ADTimestamp(dirEnt, "someProperty");
        /// DateTime wantedTime = lolwut.AsDateTime;
        /// AsDateTime will be Null if IsSetToExpire = False (prevents massive DateTime exceptions)
        /// </summary>
        public class ADTimestamp
        {
            public DateTime? AsDateTime;
            public Int64 Timestamp;
            public bool IsSetToExpire;

            public ADTimestamp(DirectoryEntry dirEnt, string propName)
            {
                this.Timestamp = GetProptInt64(dirEnt, propName);
                this.IsSetToExpire = IsStampSetToExpire(Timestamp);
                if (this.IsSetToExpire)
                {//Only calclate DateTime object if account actually has an expiration date
                    this.AsDateTime = DateTime.FromFileTime(Timestamp);
                }
            }

            /// <summary>
            /// Gets a property as an Int64 value. Mainly used with AD timestamps
            /// </summary>
            /// <param name="ent">DirEnt to get value from</param>
            /// <param name="propName">Property / Attribute to read</param>
            /// <returns>Int64 value of property</returns>
            public static Int64 GetProptInt64(DirectoryEntry ent, string propName)
            {
                //we will use the marshaling behavior of the searcher

                DirectorySearcher ds = new DirectorySearcher(ent, String.Format("({0}=*)", propName), new string[] { propName }, SearchScope.Base);

                SearchResult sr = ds.FindOne();

                if (sr != null)
                {
                    if (sr.Properties.Contains(propName))
                    {
                        return (Int64)sr.Properties[propName][0];
                    }
                }

                return -1;
            }

            public static bool IsStampSetToExpire(Int64 timestamp)
            {
                return !((timestamp == 9223372036854775807) || (timestamp == 0));
            }


        }

    }
}
