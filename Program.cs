// rmarsot@gmail.com

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// File exist... 
using System.IO;
// Ntfs ACL
using System.Security.AccessControl;
using System.Security.Principal;
using System.Security.Authentication;
using ProcessPrivileges;
// System Process 
using System.Diagnostics;
// XLS, DataRow 
using System.Data.OleDb;
using System.Data;
// LDap
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Management;
// RegEx
using System.Text.RegularExpressions;
// ReadOnlyCollection
using System.Collections.ObjectModel;

namespace FIM
{
    class Error_Handler
    {
        private struct Error_Info
        {
            private string _Description;
            private int _Level;

            public void Description(string Str) { _Description = Str; }
            public void Level(int i) { _Level = i; }
            public string Descripton() { return _Description; }
            public int Level() { return _Level; }

        }

        private Error_Info[] _Errors;
        public Error_Handler()
        {
            _Errors = new Error_Info[0];
        }

        public void Add(string Error_Description, int Error_Level)
        {
            Array.Resize(ref _Errors, _Errors.Length + 1);
            _Errors[_Errors.Length - 1].Description(Error_Description);
            _Errors[_Errors.Length - 1].Level(Error_Level);
        }

        public void Reset()
        {
            _Errors = new Error_Info[0];
        }

        public int Level()
        {
            int Max_Level = 0;
            for (int i = 0; i < _Errors.Length; i++)
            {
                if (_Errors[i].Level() > Max_Level) Max_Level = _Errors[i].Level();
            }
            return Max_Level;
        }
    }

    class Log_Type
    {
        public const int Title_1 = 101;
        public const int Title_2 = 102;
        public const int Title_3 = 103;
        public const int Title_4 = 104;
        public const int CmdLog = 10;
        public const int Error = 5;
        public const int Info = 3;
        public const int Warn = 2;
        public const int Text = 1;

        public static String HtmlCode(int logType_id, String LogText)
        {
            if (logType_id == 101) return String.Format("<h1>{0}</h1>", LogText);
            if (logType_id == 102) return String.Format("<h2>{0}</h2>", LogText);
            if (logType_id == 103) return String.Format("<h3>{0}</h3>", LogText);
            if (logType_id == 104) return String.Format("<h4>{0}</h4>", LogText);
            if (logType_id ==  10) return String.Format("<p.CmdLine>{0}</p>", LogText);
            if (logType_id ==   5) return String.Format("<p.Error>{0}</p>", LogText);
            if (logType_id ==   3) return String.Format("<p.Info>{0}</p>", LogText);
            if (logType_id ==   2) return String.Format("<p.Warn>{0}</p>", LogText);
            if (logType_id ==   1) return String.Format("<p.Text>{0}</p>", LogText);

            return "";
        }
    }

    class Log_Handler
    {
        private struct Log
        {
            private string _Description;
            private int _Level;

            public void Description(string Str) { _Description = Str; }
            public void Level(int i) { _Level = i; }
            public string Descripton() { return _Description; }
            public int Level() { return _Level; }
        }

        // private Error_Info[] _Errors;
        private Log[] _Logs;
                
        public Log_Handler()
        {
            _Logs = new Log[0];
        }
       
        public String Html()
        {
            String sHeader = @"<!DOCTYPE html PUBLIC "" -//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"" >
                                <html xmlns =""http://www.w3.org/1999/xhtml"" >
                                <head>  
                                    <meta http - equiv =""Content -Type"" content =""text /html; charset=utf-8"" />       
                                    <title> ";

            String Style = @"<style type=""text/css"">
                                        h3 {
                                            color: #000;
                                            border - bottom: 1px solid #000;
                                            padding: 5px;
                                        }

                                        h4 {
                                            margin - left: 10px;
                                            margin - top: 8px;
                                            margin - bottom: 5px;
                                        }

                                        p {
                                            margin - left: 30px;
                                            margin - top: 0px;
                                            margin - bottom: 0px;
                                        }

                                        p.info {
                                            color: #000;
	                                        background - color: #F5F5F5;
                                            border: 1px solid #C0C0C0;
                                            padding: 5px;
                                        }

                                        p.warn {
                                            color: #F60;
	                                        font - style: oblique;
                                        }

                                        p.err {
                                            color: #F00;
	                                        font - style: bold;
                                        }
                            </style>";
           
            String Html = String.Format("{0}{1}, {2}</title>{3}</head><body>",
                                        sHeader,
                                        System.Environment.MachineName,
                                         DateTime.Now.ToString("dd/MM/yyyy H:mm"), 
                                         Style);
            
            for(int i=0; i<_Logs.Length; i++)
            {
                Html += Log_Type.HtmlCode(_Logs[i].Level(), _Logs[i].Descripton());
            }

            Html += @"</body></html>";
            return Html;
        }

        public bool Html(string file_path, bool fileNamebydate = true)
        {
            StreamWriter outputFile = new StreamWriter(file_path, true);
            return true;
        }

    }

    class LDap_Operator
    {
        string _Ldap_path_begin;
        string _Ldap_path_end;
        string _ctx_ldap;
        string _ctx_path;
        DirectoryEntry _Ldap;
        PrincipalContext _Ctx;

        public LDap_Operator(string Ldap_path_begin, string Ldap_path_end, string Ctx_path_begin, string Ctx_path_end)
        {
            _Ldap = null;
            _Ctx = null;
            _Ldap_path_begin = @"LDAP://" + Ldap_path_begin + @"/";
            _Ldap_path_end = Ldap_path_end;
            _ctx_ldap = Ctx_path_begin;
            _ctx_path = Ctx_path_end;
        }

        public DirectoryEntry directoryEntry {
            get {
                if (_Ldap == null) LDap_connection_init();
                return _Ldap;
            }
        }
        public PrincipalContext principalContext {
            get {
                if (_Ctx == null) init_PrincipalContext();
                return _Ctx;
            }
        }


        public bool CheckLdapParent(string ou_parent_path = "")
        {            
            if (!LDap_connection_init(ou_parent_path)) return false;
            try
            {
                if ( (ou_parent_path.Length > 0) && (_Ldap_path_end.Length > 0))
                    return DirectoryEntry.Exists(_Ldap_path_begin + ou_parent_path + "," + _Ldap_path_end);
                if (ou_parent_path.Length > 0)
                    return DirectoryEntry.Exists(_Ldap_path_begin + ou_parent_path);
                return DirectoryEntry.Exists(_Ldap_path_begin + _Ldap_path_end);
            }
            catch(Exception e)
            {
                return false;
            }
        }

        public bool ChildAdd(OU_descriptor OU_Child)
        {
            if( !LDap_connection_init(OU_Child.Parent()) ) return false;
            if (!CheckLdapParent(OU_Child.Parent())) return false;
            try {
                DirectoryEntry newOU = _Ldap.Children.Add("OU=" + OU_Child.OU_index(0), "OrganizationalUnit");
                newOU.CommitChanges();
                newOU.Dispose();
            }
            catch (Exception e){
                return false;
            }
            
            return true;
        }


        public bool ChildrenAdd(OU_descriptor OU_Child)
        {
            if (!LDap_connection_init()) return false;
            // if (!CheckLdapParent(OU_Child.Parent())) return false;
            bool ans = true;
            string ou_descr = "";
            for (int i=0; i<OU_Child.Level(); i++)
            {                
                if(ou_descr.Length<1)
                    ou_descr = "OU=" + OU_Child.OU_index_right(i);
                else
                    ou_descr = "OU=" + OU_Child.OU_index_right(i) + "," + ou_descr;
                OU_descriptor o = new OU_descriptor(ou_descr);                
                if (!CheckLdapParent(ou_descr))
                    ans = ans & this.ChildAdd(o);
            }

            return ans;
        }

        public bool GroupAdd(string GroupName, string OU_Parent="")
        {
            if (GroupName.Length<1) return false;
            if (!LDap_connection_init(OU_Parent)) return false;
            if (Group_Check(GroupName)) return false;

            try
            {
                DirectoryEntry newGroup = _Ldap.Children.Add("CN=" + GroupName, "group");
                newGroup.CommitChanges();
                newGroup.Dispose();
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }        

        public bool Group_Check(string Group_Name)
        {

            if (!init_PrincipalContext()) return false;
            GroupPrincipal gPr = new GroupPrincipal(_Ctx, Group_Name);
            PrincipalSearcher ps = new PrincipalSearcher(gPr);
            if (ps.FindAll().Count() != 1)
            {
                ps.Dispose();
                gPr.Dispose();               
                return false;
            }

            ps.Dispose();
            gPr.Dispose();

            return true;
        }

        // Users
        public ReadOnlyCollection<string> Users(/* string GroupNameLike=""*/ )
        {
            List<string> u = new List<string>();
            if (!init_PrincipalContext()) return new ReadOnlyCollection<string>(u);                      
            UserPrincipal Users = new UserPrincipal(_Ctx);
            Users.Name = "*";
            PrincipalSearcher ps = new PrincipalSearcher(Users);
            u.Clear();
            foreach (UserPrincipal user in ps.FindAll())
            {
                u.Insert(0, user.SamAccountName);
            }            
            Users.Dispose();
            ps.Dispose();
            return new ReadOnlyCollection<string>(u);
        }

        public bool Group_add_User(string GroupName, string UserSamAccountName)
        {
            if (!init_PrincipalContext()) return false;
            UserPrincipal Users = new UserPrincipal(_Ctx);
            Users.SamAccountName = UserSamAccountName;
            PrincipalSearcher ps = new PrincipalSearcher(Users);
            if (ps.FindAll().Count() != 1) return false;
            GroupPrincipal gPr = GroupPrincipal.FindByIdentity(_Ctx,
                                                               System.DirectoryServices.AccountManagement.IdentityType.Name,
                                                               GroupName);
            foreach (var user in ps.FindAll())
            {
                try
                {                    
                    gPr.Members.Add(user);
                    gPr.Save();
                    gPr.Dispose();
                }
                catch (Exception e)
                {
                    return false;
                }
            }
            return true;
        }
       
        public bool User_Remove_from_Group(string UserName, Regex GrpNameLike /* @"(?i)^GG_STUDENT+"  */)
        {
            bool user_removed_from_group = false;
            // List<GroupPrincipal> result = new List<GroupPrincipal>();
            if (!init_PrincipalContext()) return false;
            UserPrincipal user = UserPrincipal.FindByIdentity(_Ctx, UserName);
            if (user == null) return false;            
            PrincipalSearchResult<Principal> groups = user.GetAuthorizationGroups();
            foreach (Principal p in groups)
            {
              if (p is GroupPrincipal)
              {
                    GroupPrincipal gPr = (GroupPrincipal) p;
                    if (GrpNameLike.Match(p.Name).Success)
                    {
                        gPr.Members.Remove(user);
                        gPr.Save();
                        gPr.Dispose();
                        user_removed_from_group = true;
                    }
                    //result.Add((GroupPrincipal)p);
              }
            }            
            return user_removed_from_group;
        }

        public bool User_Print_Groups(string UserNameLike)
        {
            if (!init_PrincipalContext()) return false;
            User_descriptor u = new User_descriptor(null, this._Ctx);
            if (!u.FindByIdentity(UserNameLike)) return false;
            if (!u.FindGroups()) return false;

            Console.WriteLine(" {0} is member of:", UserNameLike);
            foreach (string g in u.Groups())
            {
                Console.WriteLine("    - {0}", g);
            }

            return true;
        }


        public bool is_User_Member_of(string UserNameLike, string GroupNameLike)
        {
            if (!init_PrincipalContext()) return false;
            User_descriptor u = new User_descriptor(null, this._Ctx);
            if (!u.FindByIdentity(UserNameLike)) return false;
            if (!u.FindGroups()) return false;

            if(u.is_member_of(GroupNameLike))
                Console.WriteLine(" {0} is member of {1}", UserNameLike, GroupNameLike);
            else
                Console.WriteLine(" {0} is not a member of {1}", UserNameLike, GroupNameLike);

            return true;
        }
        
        private bool LDap_connection_init(string ou_parent_path = "")
        {
            if (_Ldap != null) _Ldap.Dispose();
            if (_Ldap_path_begin.Length < 1) return false;
            try
            {
                if (ou_parent_path.Length > 0)
                {
                    _Ldap = new DirectoryEntry(_Ldap_path_begin + ou_parent_path + "," + _Ldap_path_end);//, ADACCOUNT, ADPassword);
                }
                else
                {
                    _Ldap = new DirectoryEntry(_Ldap_path_begin + _Ldap_path_end); //, ADACCOUNT, ADPassword);                   
                }
                if(_Ldap!=null) return true;
                return false;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private bool init_PrincipalContext()
        {
            if (_Ctx != null) _Ctx.Dispose();
            if (_ctx_ldap.Length < 1)
            {
                _Ctx = new PrincipalContext(ContextType.Domain);
                if (_Ctx != null) return true;
                return false;
            }

            if (_ctx_path.Length < 1)
            {
                _Ctx = new PrincipalContext(ContextType.Domain, _ctx_ldap);
                if (_Ctx != null) return true;
                return false;
            }

            _Ctx = new PrincipalContext(ContextType.Domain, _ctx_ldap, _ctx_path);
            if (_Ctx != null) return true;
            return false;
        }
    }


    class User_descriptor
    {
        private UserPrincipal _UserPrincipal = null;
        private PrincipalContext _Context = null;
        private string[] _Groups;

        
        public User_descriptor(UserPrincipal user=null, PrincipalContext ctx=null)
        {
            this._UserPrincipal = user;
            this._Context = ctx;
        }

        public UserPrincipal User{
            get { return _UserPrincipal;  }
        }

        public bool FindByIdentity(string UserNameLike, int IdentityType=2)
        {
            // System.DirectoryServices.AccountManagement.IdentityType.SamAccountName    0 
            // System.DirectoryServices.AccountManagement.IdentityType.Name              1
            // System.DirectoryServices.AccountManagement.IdentityType.UserPrincipalName 2
            // System.DirectoryServices.AccountManagement.IdentityType.DistinguishedName 3
            // System.DirectoryServices.AccountManagement.IdentityType.Sid               4
            // System.DirectoryServices.AccountManagement.IdentityType.Guid              5

            if (_Context == null) return false;
            _UserPrincipal = null;

            if ( IdentityType==0 )
            {
                _UserPrincipal = UserPrincipal.FindByIdentity(_Context, System.DirectoryServices.AccountManagement.IdentityType.SamAccountName, UserNameLike);
                if (_UserPrincipal != null) return true;
            }

            if ( IdentityType==1 )
            {
                _UserPrincipal = UserPrincipal.FindByIdentity(_Context, System.DirectoryServices.AccountManagement.IdentityType.Name, UserNameLike);
                if (_UserPrincipal != null) return true;
            }

            if ( IdentityType==2 )
            {
                _UserPrincipal = UserPrincipal.FindByIdentity(_Context, System.DirectoryServices.AccountManagement.IdentityType.UserPrincipalName, UserNameLike);
                if (_UserPrincipal != null) return true;
            }

            if ( IdentityType==3 )
            {
                _UserPrincipal = UserPrincipal.FindByIdentity(_Context, System.DirectoryServices.AccountManagement.IdentityType.DistinguishedName, UserNameLike);
                if(_UserPrincipal != null) return true;
            }

            if ( IdentityType == 4 )
            {
                _UserPrincipal = UserPrincipal.FindByIdentity(_Context, System.DirectoryServices.AccountManagement.IdentityType.Sid, UserNameLike);
                if (_UserPrincipal != null) return true;
            }

            if ( IdentityType == 5 )
            {
                _UserPrincipal = UserPrincipal.FindByIdentity(_Context, System.DirectoryServices.AccountManagement.IdentityType.Guid, UserNameLike);
                if (_UserPrincipal != null) return true;
            }

            return false;
        }

        public bool FindGroups()
        {
            if (_Context == null) return false;
            if (_UserPrincipal == null) return false;

            PrincipalSearchResult<Principal> groups = _UserPrincipal.GetAuthorizationGroups();

            int arr_size = 0;
            foreach (Principal p in groups)
            {
                if (p is GroupPrincipal) arr_size++;
            }

            _Groups = new string[arr_size--];
            foreach (Principal p in groups)
            {
                if (p is GroupPrincipal) _Groups[arr_size--] = p.Name;
            }
          
            return true;
        }

        public bool is_member_of(string group_name)
        {
            if (_Groups == null) FindGroups();
            if (_Groups == null) return false;
            if (_Groups.Length < 1) return false;
            foreach(string g in _Groups)
            {
                if (g.ToUpper().Equals(group_name.ToUpper())) return true;
            }
            return false;
        }

        public string OU_path_to_group(int level=0)
        {
            if (_Context == null) return "";
            if (_UserPrincipal == null) return "";

            String[] substrings = _UserPrincipal.DistinguishedName.Split(',');
            List<string> ou_path = new List<string>();
            ou_path.Clear();            
            foreach (string s in substrings)
            {
                if(s.ToUpper().StartsWith("OU=") ) ou_path.Insert(0, s.ToUpper().Replace("OU=", "").Trim());
            }

            if (level > ou_path.Count) return "";
            if(level<1) return String.Join("_", ou_path.ToArray());           
            return String.Join("_", ou_path.Take(level).ToArray());
        }


        public ReadOnlyCollection<string> Groups()
        {
            List<string> s = new List<string>();
            foreach (string g in _Groups)
            {
                s.Add(g);                
            }
            return new ReadOnlyCollection<string>(s);
        }


        // System.DirectoryServices.AccountManagement.IdentityType.SamAccountName    0 
        // System.DirectoryServices.AccountManagement.IdentityType.Name              1
        // System.DirectoryServices.AccountManagement.IdentityType.UserPrincipalName 2
        // System.DirectoryServices.AccountManagement.IdentityType.DistinguishedName 3
        // System.DirectoryServices.AccountManagement.IdentityType.Sid               4
        // System.DirectoryServices.AccountManagement.IdentityType.Guid              5
        public string DistinguishedName
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";

                return _UserPrincipal.DistinguishedName;
            }
        }

        public string SamAccountName
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";

                return _UserPrincipal.SamAccountName;
            }
        }

        public string Name
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";

                return _UserPrincipal.Name;
            }
        }

        public string UserPrincipalName
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";

                return _UserPrincipal.UserPrincipalName;
            }
        }

        public string Sid
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";

                return _UserPrincipal.Sid.ToString();
            }
        }

        public string Guid
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";

                return _UserPrincipal.Guid.ToString();
            }
        }

        /*
        public string CanonicalName()
        {
            if (_Context == null) return "";
            if (_UserPrincipal == null) return "";

            var de = new DirectoryEntry(DistinguishedName());
            de.RefreshCache(new string[] { "canonicalName" });
            return de.Properties["canonicalName"].ToString();            
        }
        */

        // user.HomeDirectory
        // HomeDrive = @"U:";

        public string HomeDirectory
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";

                return _UserPrincipal.HomeDirectory;
            }

            set
            {
                if (_Context == null) return;
                if (_UserPrincipal == null) return;
                if (value.Length < 1) value = null;
                try
                {
                    _UserPrincipal.HomeDirectory = value;
                    _UserPrincipal.Save();
                }catch(Exception e)
                {
                    return;
                }
            }
        }

        public string HomeDirectory_ROOT
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";
                try
                {
                    int length = _UserPrincipal.HomeDirectory.ToLower().IndexOf(_UserPrincipal.SamAccountName.ToLower());
                    length += _UserPrincipal.SamAccountName.Length;
                    return _UserPrincipal.HomeDirectory.Substring(0, length);
                }
                catch (Exception e)
                {
                    return "";
                }
            }           
        }

        public string HomeDirectory_SRV_ROOT
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";
                try
                {
                    int length = _UserPrincipal.HomeDirectory.ToLower().IndexOf(_UserPrincipal.SamAccountName.ToLower());
                    return _UserPrincipal.HomeDirectory.Substring(0, length);
                } catch(Exception e)
                {
                    return "";
                }
            }
        }

        public string HomeDrive
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";
                try
                {
                    if (_UserPrincipal.HomeDrive == null) return "";
                    return _UserPrincipal.HomeDrive;
                }
                catch (Exception e)
                {
                    return "";
                }
            }

            set
            {
                if (_Context == null) return;
                if (_UserPrincipal == null) return;
                if (value.Length < 1) value = null;
                try
                {
                    _UserPrincipal.HomeDrive = value;
                    _UserPrincipal.Save();
                } catch(Exception e)
                {
                    return;
                }
            }
        }

        public string Script
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";

                return _UserPrincipal.ScriptPath;
            }

            set
            {
                if(value.Length<1)
                    _UserPrincipal.ScriptPath = null;
                else
                    _UserPrincipal.ScriptPath = value;
                _UserPrincipal.Save();
            }
        }

        public bool ExpirationDate_is_set()
        {
            if (_Context == null) return false;
            if (_UserPrincipal == null) return false;
            try
            {
                if(_UserPrincipal.AccountExpirationDate != null) return true;               
            }
            catch (Exception e)
            {
                return false;
            }
            return false;
        }

        public DateTime ExpirationDate
        {
            get
            {
                if (_Context == null) return DateTime.MaxValue;
                if (_UserPrincipal == null) return DateTime.MaxValue;
                try
                {
                    return _UserPrincipal.AccountExpirationDate.Value.ToLocalTime();
                }
                catch(Exception e)
                {
                    return DateTime.MaxValue;
                }
            }

            set
            {
                if (_Context == null) return;
                if (_UserPrincipal == null) return;
                try
                {
                    _UserPrincipal.AccountExpirationDate = value;
                    _UserPrincipal.Save();
                }
                catch (Exception e)
                {
                    return;
                }
            }

        }

        public bool ExpirationDate_Unset()
        {            
            if (_Context == null) return false;
            if (_UserPrincipal == null) return false;
            try
            {
                _UserPrincipal.AccountExpirationDate = null;
                _UserPrincipal.Save();
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
            

        }

        public DateTime LastLogon
        {            
            get
            {
                DateTime lastLogon = DateTime.MinValue;
                // if (_Context == null) return DateTime.MinValue; 
                // if (_UserPrincipal == null) return DateTime.MinValue;
                if (_Context == null) return lastLogon;                
                if (_UserPrincipal == null) return lastLogon;
                try
                {
                    if(_UserPrincipal.LastLogon != null ) return ((DateTime) _UserPrincipal.LastLogon);
                    return lastLogon;
                }
                catch (Exception e)
                {
                    return lastLogon;
                }
            }           
        }
      
        public DateTime LastBadPasswordAttempt
        {
            get
            {
                DateTime lastLogon = DateTime.MinValue;
                // if (_Context == null) return DateTime.MinValue; 
                // if (_UserPrincipal == null) return DateTime.MinValue;
                if (_Context == null) return lastLogon;
                if (_UserPrincipal == null) return lastLogon;
                try
                {
                    if (_UserPrincipal.LastLogon != null) return ((DateTime)_UserPrincipal.LastBadPasswordAttempt);
                    return lastLogon;
                }
                catch (Exception e)
                {
                    return lastLogon;
                }
            }
        }

        public bool AccountEnabled
        {
            get
            {
                if (_Context == null) return false;
                if (_UserPrincipal == null) return false;
                try
                {
                    return (bool) _UserPrincipal.Enabled;
                }
                catch (Exception e)
                {
                    return false;
                }
            }

            set
            {
                if (_Context == null) return;
                if (_UserPrincipal == null) return;
                try
                {
                    _UserPrincipal.Enabled = value;
                    _UserPrincipal.Save();
                }
                catch (Exception e)
                {
                    return;
                }
            }

        }

        public bool Account_Expirated()
        {
            if (DateTime.Compare(ExpirationDate, DateTime.Now) > 0) return false;
            return true;
        }

        public string Description
        {
            get
            {
                if (_Context == null) return "";
                if (_UserPrincipal == null) return "";

                var UserEntry = (DirectoryEntry)_UserPrincipal.GSTUDENTnderlyingObject();
                if (UserEntry.Properties["Description"].Value == null) return "";
                return UserEntry.Properties["Description"].Value.ToString();
            }

            set
            {
                if (_Context == null) return;
                if (_UserPrincipal == null) return;
                try
                {
                    var UserEntry = (DirectoryEntry)_UserPrincipal.GSTUDENTnderlyingObject();
                    UserEntry.Properties["Description"].Value = value;
                    UserEntry.CommitChanges();
                }catch(Exception e)
                {
                    return;
                }
            }
        }

        public bool CreateHome(string HomePath)
        {
            if (_UserPrincipal == null) return false;
            if (Directory.Exists(HomePath + @"\" + SamAccountName)) return false;
            UserFolders uF = new UserFolders();
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName, _UserPrincipal.SamAccountName, 1)) return false;
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName + @"\DATA", _UserPrincipal.SamAccountName, 1)) return false;
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName + @"\Profil", _UserPrincipal.SamAccountName, 2)) return false;
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName + @"\Profil.V2", _UserPrincipal.SamAccountName, 2)) return false;
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName + @"\Profil.TSE", _UserPrincipal.SamAccountName, 2)) return false;
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName + @"\DATA\Musique", _UserPrincipal.SamAccountName, 2)) return false;
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName + @"\DATA\Vidéos", _UserPrincipal.SamAccountName, 2)) return false;
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName + @"\DATA\Favoris", _UserPrincipal.SamAccountName, 2)) return false;
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName + @"\DATA\Images", _UserPrincipal.SamAccountName, 2)) return false;
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName + @"\DATA\Bureau", _UserPrincipal.SamAccountName, 2)) return false;
            if (!uF.CreateDirectory(HomePath + @"\" + SamAccountName + @"\DATA\Documents", _UserPrincipal.SamAccountName, 2)) return false;

            return true;
        }

        public bool MoveHome(string old_HomePath, string new_HomePath)
        {
            if (old_HomePath.Length<1) return false;
            if (new_HomePath.Length<1) return false;
            if (_UserPrincipal == null) return false;            
            UserFolders uF = new UserFolders();                                    
            if (!old_HomePath.EndsWith(@"\")) old_HomePath += @"\";
            if (!new_HomePath.EndsWith(@"\")) new_HomePath += @"\";
            if (!uF.CreateDirectory(new_HomePath + SamAccountName)) return false;
            if (!uF.Copy(old_HomePath + SamAccountName + @"\", new_HomePath + SamAccountName)) return false;
            if (!uF.DeleteDirectory(old_HomePath + SamAccountName)) return false;            
            
            return true;
        }

        public bool Move_To_OU(string newOU)
        {            
            if (_UserPrincipal == null) return false;

            // Console.WriteLine(_UserPrincipal.DistinguishedName);
            // Console.WriteLine(newOU);

            DirectoryEntry eLocation = new DirectoryEntry("LDAP://" + _UserPrincipal.DistinguishedName);
            DirectoryEntry nLocation = new DirectoryEntry("LDAP://" + newOU);
            try
            {
                eLocation.MoveTo(nLocation);
                nLocation.Close();
                eLocation.Close();
            }catch
            {
                return false;
            }

            return true;
        }        
    }

    /// <summary>
    /// Class: UserFolders
    /// </summary>
    class UserFolders
    {        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Path">Directory path</param>
        /// <param name="SamAccountName">SamAccountName</param>
        /// <param name="User_ACL_rights">0 -> Not Set, 1 -> ReadOnly, 2 -> Allow SubFolders write access</param>
        /// <returns></returns>
        public bool CreateDirectory(string Path, string SamAccountName="", int User_ACL_rights=0)
        {
            try
            {
                // Create Directories
                Directory.CreateDirectory(Path);               
                if (!Directory.Exists(Path)) return false;
                switch (User_ACL_rights)
                {
                    case 1:
                        if (!Folder_User_ReadOnly_ACL(Path, SamAccountName)) return false;
                        break;
                    case 2:
                        if(!sub_Folder_User_Write_ACL(Path, SamAccountName)) return false;
                        break;
                    default:                        
                        break;
                }
            }
            catch (Exception e)
            {
                return false;
            }

            return true;
        }

        public bool DeleteDirectory(string dir_path)
        {
            try
            {              
                Directory.Delete(dir_path, true);
            }
            catch (Exception e)
            {
                return false;
            }

            return true;

        }

        public bool MoveDirectory(string old_dir_path, string new_dir_path)
        {
            try
            {
                if (!Directory.Exists(old_dir_path)) return false;
                if (Directory.Exists(new_dir_path)) return false;
                Directory.Move(old_dir_path, new_dir_path);
                
                if (Directory.Exists(old_dir_path)) return false;
                if (Directory.Exists(new_dir_path)) return true;
            } catch(Exception e)
            {
                return false;
            }

            return false;
        }

        public bool Copy(string src, string dest, UserPrincipal in_Owner = null)
        {
            if (!Directory.Exists(src)) return false;
            if (!Directory.Exists(dest)) return false;
            DirectoryCopy(src, dest, true, in_Owner);

            return true;
        }


        /* --------------------------------------------- 
            Directory copy
        ------------------------------------------------ */
        private void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs, UserPrincipal in_Owner = null)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Console.Write("[D] Create Directory {0}", destDirName);
                try
                {
                    Directory.CreateDirectory(destDirName);
                    Console.WriteLine(" [OK]");
                }
                catch
                {
                    Console.WriteLine(" [FAIL]");
                }
                if (in_Owner != null)
                {
                    if (SetFileOwner(destDirName, in_Owner))
                        Console.WriteLine("  > [OK ] Set Owner {0} ({1})", in_Owner.UserPrincipalName, in_Owner.SamAccountName);
                    else
                        Console.WriteLine("  > [FAIL] Set Owner {0} ({1})", in_Owner.UserPrincipalName, in_Owner.SamAccountName);
                }
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                Console.Write("[F] Copy File {0}", temppath);
                if (temppath.ToLower().IndexOf("fritzing.exe") < 0)
                {
                    try
                    {
                        file.CopyTo(temppath, true);
                        Console.WriteLine(" [OK]");
                    }
                    catch
                    {
                        Console.WriteLine(" [FAIL]");
                    }
                }
                if (in_Owner != null)
                {
                    if (SetFileOwner(temppath, in_Owner))
                        Console.WriteLine("  > [OK ] Set Owner {0}", in_Owner);
                    else
                        Console.WriteLine("  > [FAIL] Set Owner {0}", in_Owner);
                }
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs, in_Owner);
                }
            }
        }



        // Change Owner            
        public void DirSearch_set_owner(string dir, UserPrincipal in_Owner, bool verbose = false)
        {
            try
            {
                foreach (string f in Directory.GetFiles(dir))
                {
                    if (verbose) Console.Write("{0} set owner {1}", f, in_Owner.UserPrincipalName);
                    if (SetFileOwner(f, in_Owner))
                    {
                        if (verbose) Console.WriteLine(" [OK]");
                    }
                    else
                    {
                        if (verbose) Console.WriteLine(" [FAIL]");
                    }
                }
                foreach (string d in Directory.GetDirectories(dir))
                {
                    if (verbose) Console.Write("{0} set owner {1}", d, in_Owner.UserPrincipalName);
                    if (SetFileOwner(d, in_Owner))
                    {
                        if (verbose) Console.WriteLine(" [OK]");
                    }
                    else
                    {
                        if (verbose) Console.WriteLine(" [FAIL]");
                    }
                    DirSearch_set_owner(d, in_Owner);
                }

            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public bool SetFileOwner(string in_filePath, UserPrincipal in_OwnNer)
        {
            try
            {
                using (new ProcessPrivileges.PrivilegeEnabler(Process.GetCurrentProcess(), Privilege.TakeOwnership))
                {
                    DirectoryInfo dInfo = new DirectoryInfo(in_filePath);
                    DirectorySecurity dSec = dInfo.GetAccessControl();
                    dSec.SetOwner(in_OwnNer.Sid);
                    Directory.SetAccessControl(dInfo.FullName, dSec);
                }
                return true;
            }
            catch (Exception e)
            {
                // Console.WriteLine("Err {0}", e.ToString());
                return false;
            }

        }


        /// <summary>
        /// Process
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="arg"></param>
        /// <param name="in_TimeOut"></param>
        /// <returns></returns>
        private bool RunProcess_and_Wait_for_exit(string cmd, string arg, int in_TimeOut = 0)
        {
            try
            {
                ProcessStartInfo start = new ProcessStartInfo(cmd, arg);
                start.UseShellExecute = false;
                start.CreateNoWindow = false;
                Process P = Process.Start(start);
                if (in_TimeOut > 0 && P.WaitForExit(in_TimeOut) == false)
                {
                    P.Kill();
                    return false;
                }
                else
                    P.WaitForExit();
                if (P.ExitCode != 0) return false;
                return true;
            }
            catch
            {
                // Console.WriteLine("[FAIL] Process {0} with args : {1} ", cmd, arg);
                return false;
            }

        }

        /// <summary>
        /// ACL -> Readonly for user
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="SamAccountName"></param>
        /// <returns></returns>
        public bool Folder_User_ReadOnly_ACL(string folderPath, string SamAccountName)
        {
            try
            {
                bool modified;
                
                FileSystemAccessRule accessRule = new FileSystemAccessRule(identity: SamAccountName,
                                                                            fileSystemRights: FileSystemRights.Read,
                                                                            type: AccessControlType.Allow,
                                                                            inheritanceFlags: InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                                                            propagationFlags: PropagationFlags.None);

                DirectoryInfo dInfo = new DirectoryInfo(folderPath);
                DirectorySecurity dSecurity = dInfo.GetAccessControl();
                dSecurity.ModifyAccessRule(AccessControlModification.Set, accessRule, out modified);
                
                System.Security.Principal.NTAccount group = new System.Security.Principal.NTAccount("domain.net", "Administrateurs");
                FileSystemAccessRule accessRule3 = new FileSystemAccessRule(identity: group,
                                                                            fileSystemRights: FileSystemRights.FullControl,
                                                                            type: AccessControlType.Allow,
                                                                            inheritanceFlags: InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                                                            propagationFlags: PropagationFlags.None);
                dSecurity.ModifyAccessRule(AccessControlModification.Add, accessRule3, out modified);

                FileSystemAccessRule accessRule4 = new FileSystemAccessRule("Administrateurs",
                                                                            fileSystemRights: FileSystemRights.FullControl,
                                                                            type: AccessControlType.Allow,
                                                                            inheritanceFlags: InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                                                            propagationFlags: PropagationFlags.None);
                dSecurity.ModifyAccessRule(AccessControlModification.Add, accessRule4, out modified);

                dSecurity.SetAccessRuleProtection(true, false);

                dInfo.SetAccessControl(dSecurity);
            }
            catch (Exception e)
            {                
                return false;
            }
            return true;
        }

        public bool sub_Folder_User_Write_ACL(string folderPath, string SamAccountName)
        {
            bool modified;

            try
            {
                FileSystemAccessRule accessRule = new FileSystemAccessRule(identity: SamAccountName,
                                                                             fileSystemRights: FileSystemRights.FullControl,
                                                                             type: AccessControlType.Allow,
                                                                             inheritanceFlags: InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                                                             propagationFlags: PropagationFlags.InheritOnly);
                DirectoryInfo dInfo = new DirectoryInfo(folderPath);
                DirectorySecurity dSecurity = dInfo.GetAccessControl();
                dSecurity.ModifyAccessRule(AccessControlModification.Set, accessRule, out modified);

                FileSystemAccessRule accessRule2 = new FileSystemAccessRule(identity: SamAccountName,
                                                                            fileSystemRights: FileSystemRights.Read | FileSystemRights.Write | FileSystemRights.Traverse,
                                                                            type: AccessControlType.Allow,
                                                                            inheritanceFlags: InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                                                            propagationFlags: PropagationFlags.None);
                dSecurity.ModifyAccessRule(AccessControlModification.Add, accessRule2, out modified);


                System.Security.Principal.NTAccount group = new System.Security.Principal.NTAccount("domain.net", "Administrateurs");
                FileSystemAccessRule accessRule3 = new FileSystemAccessRule(identity: group,
                                                                            fileSystemRights: FileSystemRights.FullControl,
                                                                            type: AccessControlType.Allow,
                                                                            inheritanceFlags: InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                                                            propagationFlags: PropagationFlags.None);
                dSecurity.ModifyAccessRule(AccessControlModification.Add, accessRule3, out modified);

                FileSystemAccessRule accessRule4 = new FileSystemAccessRule("Administrateurs",
                                                                            fileSystemRights: FileSystemRights.FullControl,
                                                                            type: AccessControlType.Allow,
                                                                            inheritanceFlags: InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                                                            propagationFlags: PropagationFlags.None);
                dSecurity.ModifyAccessRule(AccessControlModification.Add, accessRule4, out modified);

                dSecurity.SetAccessRuleProtection(true, false);

                dInfo.SetAccessControl(dSecurity);
            }
            catch (Exception e)
            {
                return false;
            }

            return true;
        }

        static bool TakeOwnerShip(string in_Path, IdentityReference in_User /* WindowsIdentity.GetCurrent().User */, bool PrintError = false)
        {
            try
            {
                using (new ProcessPrivileges.PrivilegeEnabler(Process.GetCurrentProcess(), Privilege.TakeOwnership))
                {
                    DirectoryInfo dInfo = new DirectoryInfo(in_Path);
                    DirectorySecurity dSec = dInfo.GetAccessControl();                    
                    dSec.SetOwner(WindowsIdentity.GetCurrent().User);
                    Directory.SetAccessControl(dInfo.FullName, dSec);
                }
            }
            catch (Exception e)
            {
                if (PrintError) Console.WriteLine(e.Message);
                return false;
            }

            return true;
        }

    }

    class OU_descriptor
    {
        // private DataRow _RAW_Data_Line;
        private string[] _Level_info;
        private string _Code_Etape = "";

        public OU_descriptor(DataRow RAW_dtR, string CodeEtape="")
        {
            _Level_info = new string[0];
            if(CodeEtape.Length<12)
                _Code_Etape = CodeEtape;
            else
                _Code_Etape = "";
            if (RAW_dtR.ItemArray.Length > 2)
            {
                int size_with_out_empty = 0;
                for (int i = 0; i < RAW_dtR.ItemArray.Length - 2; i++)
                {
                    if (RAW_dtR[i + 2].ToString().Length > 1) size_with_out_empty++;
                }
                _Level_info = new string[size_with_out_empty];
                int dt_pos = 0;
                for (int i = 0; i < RAW_dtR.ItemArray.Length - 2; i++)
                {
                    if (RAW_dtR[i + 2].ToString().Length > 1)
                        _Level_info[dt_pos++] = RAW_dtR[i + 2].ToString();
                }
            }

        }

        public OU_descriptor(string ou_descriptor)
        {
            _Level_info = ou_descriptor.Split(',');
            for(int i=0; i<_Level_info.Length; i++)
            {
                _Level_info[i] = _Level_info[i].Replace("OU=", "").Trim();
                _Level_info[i] = _Level_info[i].Replace("ou=", "").Trim();
                _Level_info[i] = _Level_info[i].Replace("Ou=", "").Trim();
                _Level_info[i] = _Level_info[i].Replace("oU=", "").Trim();
            }
        }

        public string Code_Etape() { return _Code_Etape; }

        public string OU(int MaxLevel = 0)
        {
            if (this.Level() < 1) return "";
            if (MaxLevel < 1) MaxLevel= this.Level();
            string ans = "";
            if (MaxLevel > this.Level()) MaxLevel = this.Level();
            for (int i = 0; i < MaxLevel; i++)
            {
                if (ans.Length > 0) ans = "," + ans;
                ans = "OU=" + _Level_info[i] + ans;
            }
            return ans;
        }

        public string Parent()
        {
            if (this.Level() < 2) return "";            
            string ans = "";            
            for (int i = this.Level() - 1; i>0; i--)
            {
                if (ans.Length > 0) ans = "," + ans;
                ans = "OU=" + _Level_info[i] + ans;
            }
            return ans;
        }

        public string OU_reverse(int MaxLevel = 0)
        {
            if (this.Level() < 1) return "";
            if (MaxLevel < 1) MaxLevel = this.Level();
            string ans = "";
            if (MaxLevel > this.Level()) MaxLevel = this.Level();
            for (int i = 0; i < MaxLevel; i++)
            {
                if (ans.Length > 0) ans = "," + ans;
                ans = "OU=" + _Level_info[this.Level()-1-i] + ans;
            }
            return ans;
        }

        public string OU_index(int Level)
        {
            if (this.Level() < 1) return "";
            if (Level < 0) return "";           
            if (Level >= this.Level()) Level = this.Level()-1;
            return _Level_info[Level];
        }

        public string OU_index_right(int Level)
        {
            if (this.Level() < 1) return "";
            if (Level < 0) return "";
            if (Level >= this.Level()) Level = this.Level()-1;
            return _Level_info[this.Level()-Level-1];
        }

        public string OU_Group(int MaxLevel = 0)
        {
            if (this.Level() < 1) return "";
            if (MaxLevel < 1) MaxLevel = this.Level();
            string ans = "";            
            if (MaxLevel > this.Level()) MaxLevel = this.Level();
            for (int i = 0; i < MaxLevel; i++)
            {
                if (ans.Length > 0) ans += "_";
                ans += _Level_info[i];
            }
            return ans;
        }

        public int Level()
        {
            return _Level_info.Length;
        }
    }

    class Structure_Xlsx_Config_Loader
    {
        Error_Handler Err = new Error_Handler();
        DataSet xls_Dataset;
        OU_descriptor[] OU_List; 

        public Structure_Xlsx_Config_Loader(string FileName)
        {
            // init 
            xls_Dataset = new DataSet();
            OU_List = new OU_descriptor[0];
            LoadFile(FileName);
        }

        public void print_XML()
        {
            Console.WriteLine("{0}", this.XML());
        }

        public void print_OU()
        {
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STUDENT,OU=local,DC=domain,DC=fr", 
                                                 "domain.net", "");            

            foreach (string o in OrganizationalUnits_Names())
            {
                string oExist = "[NOT FOUND]";                
                if (op.CheckLdapParent(o ))
                    oExist = "[OK]";
                /*
                else
                {                   
                    OU_descriptor oD = new OU_descriptor(o.ToUpper());
                    op.ChildrenAdd(oD);
                } 
                */
                    
                if(this.Code_Etape(o).Length>0)                          
                    Console.WriteLine("OU {0} {1} [{2}]", o, oExist, this.Code_Etape(o).ToUpper());
                else
                    Console.WriteLine("OU {0} {1}", o, oExist);                

            }

            //OU_descriptor OU = new OU_descriptor("OU=TEST2,OU=NEW");
            // op.ChildrenAdd(OU);

        }

        public void print_GROUPS()
        {
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STUDENT,OU=local,DC=domain,DC=fr",
                                                 "domain.net", "");            
            foreach (string s in Groups())
            {
                string gName = "GG_STUDENT_" + s;
                string gExist = "[NOT FOUND]";
                if (op.Group_Check(gName)) gExist = "[OK]";
                if (this.Code_Etape(s).Length > 0)
                    Console.WriteLine("Group {0} {1} [{2}]", gName, gExist, this.Code_Etape(s).ToUpper());
                else
                    Console.WriteLine("Group {0} {1}", gName, gExist);
            }            
        }

        public void create_GROUPS()
        {
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STUDENT,OU=local,DC=domain,DC=fr",
                                                 "domain.net", "");
            foreach (string s in Groups())
            {
                string gName = "GG_STUDENT_" + s;
                string gExist = "[NOT FOUND]";
                if (op.Group_Check(gName))
                    gExist = "[OK]";
                else
                {
                    if(op.GroupAdd(gName))
                        gExist = "[CREATED]";
                }

                if (this.Code_Etape(s).Length > 0)
                    Console.WriteLine("Group {0} {1} [{2}]", gName, gExist, this.Code_Etape(s).ToUpper());
                else
                    Console.WriteLine("Group {0} {1}", gName, gExist);
            }
        }

        public string XML()
        {
            string xml_ans = "";
            xml_ans += "<Field name=\"ou\" type=\"choice\">" + System.Environment.NewLine;
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STUDENT,OU=local,DC=domain,DC=fr",
                                                 "domain.net", "");
            for (int i = 0; i < OU_List.Length; i++)
            {
                if (OU_List[i].Code_Etape().Length>0)
                    xml_ans += String.Format("<Condition value=\"{0}\" type=\"startswith\" field=\"domainCodeVersionEtape\">{1},OU=STUDENT,OU=local,DC=domain,DC=fr</Condition >{2}",
                                              OU_List[i].Code_Etape(), OU_List[i].OU(), System.Environment.NewLine);
            }

            xml_ans += "</Field>" + System.Environment.NewLine;
            return xml_ans; 

        }

        public ReadOnlyCollection<OU_descriptor> OrganizationalUnits()
        {
            List<OU_descriptor> OUs = new List <OU_descriptor> ();
            foreach (OU_descriptor o in OU_List) OUs.Add(o);
            return new ReadOnlyCollection<OU_descriptor> (OUs);
        }

        public ReadOnlyCollection<string> OrganizationalUnits_Names()
        {
            List<string> s = new List<string>();
            foreach (OU_descriptor o in OU_List)
            {
                for (int i = 0; i < o.Level(); i++)
                {
                    if (!string_exist_in_array(s.ToArray(), o.OU(i), true))
                        s.Add(o.OU(i));
                }
            }
            return new ReadOnlyCollection<string>(s);
        }

        public string Code_Etape(string OU_description)
        {
            foreach (OU_descriptor o in OU_List)
            {
                if (o.OU().Equals(OU_description)) return o.Code_Etape();
            }
            return "";
        }

        public ReadOnlyCollection<string> Groups()
        {
            List<string> s = new List<string>();
            foreach (OU_descriptor o in OU_List)
            {
                for (int i = 0; i < o.Level(); i++)
                {
                    if( !string_exist_in_array(s.ToArray(), o.OU_Group(i), true) )
                        s.Add(o.OU_Group(i));
                }
            }
            return new ReadOnlyCollection<string>(s);
        }

        private bool string_exist_in_array(string[] sArr, string find_me, bool ignore_case=false)
        {
            if (sArr.Length < 1) return false;
            if (find_me.Length < 1) return false;
            if (ignore_case) find_me = find_me.ToLower();
            foreach(string s in sArr)
            {
                string sc = s;
                if (ignore_case) sc = s.ToLower();
                if (sc.Equals(find_me)) return true;                
            }
            return false;
        }

        private bool LoadFile(string xlsxFileName)
        {
            xls_Dataset.Clear();
            if (!File.Exists(xlsxFileName))
            {
                Err.Add("Unable to open file " + xlsxFileName, 3);
                return false;
            }

            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + xlsxFileName + @";Extended Properties=Excel 12.0 Xml";

            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [local-OU$]", connection);
                OleDbDataAdapter adap = new OleDbDataAdapter(command);
                adap.Fill(xls_Dataset);
                if (xls_Dataset.Tables.Count < 1)
                {
                    Err.Add(xlsxFileName + " sheet [local-OU] Was not found", 3);
                    return false;
                }
                if (xls_Dataset.Tables[0].Columns.Count < 2)
                {
                    Err.Add(xlsxFileName + " sheet [local-OU] was found, but has not at least 2 columns", 3);
                    return false;
                }
                if (xls_Dataset.Tables[0].Rows.Count < 4)
                {
                    Err.Add(xlsxFileName + " sheet [local-OU] was found, but seems to have less than 3 lines", 3);
                    return false;
                }

                int OU_List_size = 0;
                for (int i = 0; i < xls_Dataset.Tables[0].Rows.Count - 2; i++)
                {
                    if (xls_Dataset.Tables[0].Rows[i][0].ToString().Length > 1 || xls_Dataset.Tables[0].Rows[i][1].ToString().Length > 1)
                        OU_List_size++;
                }

                OU_List = new OU_descriptor[--OU_List_size];
                int OU_index = 0;
                for (int i = 1; i < xls_Dataset.Tables[0].Rows.Count - 2; i++)
                {
                    if(xls_Dataset.Tables[0].Rows[i][0].ToString().Length>1 || xls_Dataset.Tables[0].Rows[i][1].ToString().Length > 1)
                    OU_List[OU_index++] = new OU_descriptor(xls_Dataset.Tables[0].Rows[i], xls_Dataset.Tables[0].Rows[i][1].ToString());
                }                                                 
            }
            return true;
        }

    }






    class Program
    {
        // STUDENT 
        static void STUDENT()
        {
            // List STUDENT users from Temp.
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STUDENT,OU=local,DC=domain,DC=fr",
                                                 "domain.net", "OU=STUDENT,OU=local,DC=domain,DC=fr");

            LDap_Operator op_Global = new LDap_Operator("domain.net", "OU=local,DC=domain,DC=fr",
                                                        "domain.net", "OU=local,DC=domain,DC=fr");
            User_descriptor u;
            foreach (string userSam in op.Users())
            {
                u = new User_descriptor(null, op.principalContext);
                if (u.FindByIdentity(userSam, 0))
                {
                    u.FindGroups();
                    string MainGroup = u.OU_path_to_group().Replace("", "GG_");
                    if (u.is_member_of(MainGroup) && !u.is_member_of("COMPTES_TEMPO"))
                    {
                        Console.WriteLine("{0} (know as {1}), in {2}", u.SamAccountName, u.UserPrincipalName, u.OU_path_to_group());
                        int level = 2;
                        while (u.OU_path_to_group(level).Length > 0)
                        {
                            string Grp = u.OU_path_to_group(level++).Replace("", "GG_");                            
                            Console.WriteLine(" is member of {0} -> ok", Grp);
                        }
                    }
                   
                }
                u = null;
            }

            /*
            if (u.ExpirationDate_is_set() && !u.is_member_of("COMPTES_TEMPO"))
            {
               Console.WriteLine("{0} (know as {2}) has exp date set to {1}", userSam, u.ExpirationDate, u.UserPrincipalName);
               u.ExpirationDate_Unset();
            }
            */

            /*
           User_descriptor u;            
           foreach (string userSam in op.Users())
           {                
               u = new User_descriptor(null, op.principalContext);
               if (u.FindByIdentity(userSam, 0))
               {
                   u.FindGroups();
                   string MainGroup = u.OU_path_to_group().Replace("", "GG_");
                   if (u.is_member_of(MainGroup))
                       Console.WriteLine("{0} is member of {1}", u.Name, MainGroup);
                   else
                   {
                       Console.WriteLine("{0} is not member of {1}!", u.Name, MainGroup);
                       System.Text.RegularExpressions.Regex GrpNameLike = new Regex(@"(?i)^GG_STUDENT+");
                       op.User_Remove_from_Group(u.Name, GrpNameLike);
                       int level = 2;
                       while (u.OU_path_to_group(level).Length > 0)
                       {
                           string Grp = u.OU_path_to_group(level++).Replace("", "GG_");
                           if (op_Global.Group_add_User(Grp, u.User.SamAccountName))
                               Console.WriteLine(" added to {0} -> ok", Grp);
                           else
                               Console.WriteLine(" added to {0} -> FAIL!!!", Grp);
                       }
                   }
               }
               u = null;                               
           }

           */


        }

        static void STUDENT_archives()
        {
            // List STUDENT users from Temp.
            LDap_Operator op = new LDap_Operator("domain.net", "OU=Temp,OU=local,DC=domain,DC=fr",
                                                 "domain.net", "OU=Temp,OU=local,DC=domain,DC=fr");
            User_descriptor u;
            foreach (string userSam in op.Users())
            {
                u = new User_descriptor(null, op.principalContext);
                u.CreateHome(@"\\");
            }                
        }


        static void STAFF()
        {
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STAFF,OU=local,DC=domain,DC=fr",
                                                 "domain.net", "OU=STAFF,OU=local,DC=domain,DC=fr");

            LDap_Operator op_Global = new LDap_Operator("domain.net", "OU=local,DC=domain,DC=fr",
                                                        "domain.net", "OU=local,DC=domain,DC=fr");

            User_descriptor u;
            foreach (string userSam in op.Users())
            {
                u = new User_descriptor(null, op.principalContext);
                if (u.FindByIdentity(userSam, 0))
                {          
                    if (op_Global.Group_add_User(@"GG_Adm", userSam))
                        Console.WriteLine(@"Add user {0} ({2}) to group {1} [ OK ]", userSam, @"GG_Adm", u.UserPrincipalName);
                    else
                        Console.WriteLine(@"Add user {0} ({2}) to group {1} [FAIL]", userSam, @"GG_Adm", u.UserPrincipalName);
                    u = null;
                    // STAFF(userSam);
                }
                else
                    Console.WriteLine(@"!!! User {0} not found", userSam);
            }
            
        }

        static void STAFF(string userSam)
        {
            // List STUDENT users from Temp.
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STAFF,OU=local,DC=domain,DC=fr",
                                                 "domain.net", "OU=STAFF,OU=local,DC=domain,DC=fr");

            LDap_Operator op_Global = new LDap_Operator("domain.net", "OU=local,DC=domain,DC=fr",
                                                        "domain.net", "OU=local,DC=domain,DC=fr");

            User_descriptor u = new User_descriptor(null, op.principalContext);
            if (!u.FindByIdentity(userSam, 0)) return;
            if (u.is_member_of("GG_Adm")) return;
            u.CreateHome(@"\\localESRVDATA2\E$\Users");
            u.HomeDirectory = @"\\localESRVDATA2\Users$\" + userSam + @"\DATA\Documents";
            u.HomeDrive = @"U:";
            op_Global.Group_add_User("GG_Adm_SG", userSam);
            u.Script = @"local\Login.cmd";
        }



        // CNAM
        static void STAFF_CNAM()
        {
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STAFF,OU=CNAM,DC=domain,DC=fr",
                                                 "domain.net", "OU=STAFF,OU=CNAM,DC=domain,DC=fr");
            foreach (string userSam in op.Users())
            {
                Console.WriteLine("{0}", userSam);
                STAFF_CNAM(userSam);
            }
            
        }
        
        static void STAFF_CNAM(string userSam)
        {
            // List STUDENT users from Temp.
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STAFF,OU=CNAM,DC=domain,DC=fr",
                                                 "domain.net", "OU=STAFF,OU=CNAM,DC=domain,DC=fr");

            LDap_Operator op_Global = new LDap_Operator("domain.net", "DC=domain,DC=fr",
                                                        "domain.net", "DC=domain,DC=fr");

            User_descriptor u = new User_descriptor(null, op.principalContext);
            if (!u.FindByIdentity(userSam, 0)) return;
            if (u.is_member_of("GG_Adm_Cnam")) return;
            u.CreateHome(@"\\localESRVDATA2\E$\Users");
            u.HomeDirectory = @"\\localESRVDATA2\Users$\" + userSam + @"\DATA\Documents";
            u.HomeDrive = @"U:";
            op_Global.Group_add_User("GG_Adm_CNAM", userSam);
            u.Script = @"";

        }

        static void STAFF_Info()
        {
            // List STUDENT users from Temp.
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STAFF,OU=local,DC=domain,DC=fr",
                                                 "domain.net", "OU=STAFF,OU=local,DC=domain,DC=fr");


            User_descriptor u;
            int i = 0;
            foreach (string userSam in op.Users())
            {
                u = new User_descriptor(null, op.principalContext);
                // Move_User_to_Archives(op, userSam, @"\\localesrvdata3\e$\Archives");
                // if (u.is_member_of("GG_Adm")) return;
                
                if (u.FindByIdentity(userSam, 0))
                {
                    if (!u.is_member_of("GG_Adm") && !u.is_member_of("COMPTES_TEMPO"))
                    {
                        i++;
                        Console.WriteLine("{2} {0} ({1}) [{3}]", userSam, u.UserPrincipalName, String.Format("{0, 3}", i), u.DistinguishedName );
                        STAFF(userSam);
                    }
                    // if(u.LastBadPasswordAttempt.CompareTo(DateTime.MinValue)>0)
                    //    Console.WriteLine("{0} is {1}, last logon: {2}", userSam, u.UserPrincipalName, u.LastBadPasswordAttempt);
                }
                
                u = null;
            }
        }        


        // @"\\localesrvdata3\e$\Archives"
        static bool Move_User_to_Archives(LDap_Operator op, string userSam, string achives_path)
        {
            if (op==null) return false;
            /*
            LDap_Operator op = new LDap_Operator("domain.net", "OU=Temp,OU=local,DC=domain,DC=fr",
                                                 "domain.net", "OU=Temp,OU=local,DC=domain,DC=fr");
            */
            User_descriptor u = new User_descriptor(null, op.principalContext);                        
            if (!u.FindByIdentity(userSam, 0)) return false;
            Console.WriteLine("{0} is {1} home root is {2}", u.SamAccountName, u.Name, u.HomeDirectory_ROOT);
            if (!u.AccountEnabled)
            {
                Console.WriteLine("  --> Disabled account, nothing to do...");
                return true;
            }

            if (!u.ExpirationDate_is_set())
            {
                u.ExpirationDate = DateTime.Now.AddDays(30);
                Console.WriteLine("  --> Set ExpirationDate to {0}", u.ExpirationDate.ToString());
                return true;
            }

            if (u.Account_Expirated())
            {
                if(!op.is_User_Member_of(u.SamAccountName, "GG_Archives"))
                    op.Group_add_User("GG_Archives", u.SamAccountName);
            }


                /*
                if (u.Account_Expirated())
                {

                    Console.WriteLine("  --> Move directory to archives form {0} to {1}", 
                                        u.HomeDirectory_SRV_ROOT,
                                        achives_path);
                    u.AccountEnabled = false;
                    u.MoveHome(u.HomeDirectory_SRV_ROOT, achives_path);
                    if(u.Description.Length>0)
                        u.Description += String.Format(@" // Home: {0}\{1}", achives_path, u.SamAccountName);
                    else
                        u.Description = String.Format(@"Home: {0}\{1}", achives_path, u.SamAccountName);

                    u.AccountEnabled = 
                    u.HomeDrive = "";
                    u.HomeDirectory = ""; 


                }
                */

                return true;           
        }


        static void Archives()
        {
            LDap_Operator op = new LDap_Operator("domain.net", "OU=Temp,OU=local,DC=domain,DC=fr",
                                                 "domain.net", "OU=Temp,OU=local,DC=domain,DC=fr");

            foreach (string userSam in op.Users())
            {
                Move_User_to_Archives(op, userSam, @"\\localesrvdata3\e$\Archives");
            }
        }

        static void Main(string[] args)
        {
            LDap_Operator op = new LDap_Operator("domain.net", "DC=domain,DC=fr",
                                                "domain.net", "DC=domain,DC=fr");

            User_descriptor u = new User_descriptor(null, op.principalContext);
            u.FindByIdentity("s0700139", 0);
            Console.WriteLine("Move {0}", u.UserPrincipalName);
            if (u.Move_To_OU("OU=Logistique,OU=SG,OU=STAFF,OU=local,DC=domain,DC=fr"))
            {
                Console.WriteLine("OK");
            }
            else
            {
                Console.WriteLine("FAIL");
            }


            u.FindByIdentity("s1200225", 0);
            Console.WriteLine("Move {0}", u.UserPrincipalName);
            if (u.Move_To_OU("OU=Logistique,OU=SG,OU=STAFF,OU=local,DC=domain,DC=fr"))
            {
                Console.WriteLine("OK");
            }
            else
            {
                Console.WriteLine("FAIL");
            }


            u.FindByIdentity("s1200275", 0);
            Console.WriteLine("Move {0}", u.UserPrincipalName);
            if (u.Move_To_OU("OU=Logistique,OU=SG,OU=STAFF,OU=local,DC=domain,DC=fr"))
            {
                Console.WriteLine("OK");
            }
            else
            {
                Console.WriteLine("FAIL");
            }


            u.FindByIdentity("s0500063", 0);
            Console.WriteLine("Move {0}", u.UserPrincipalName);
            if (u.Move_To_OU("OU=Logistique,OU=SG,OU=STAFF,OU=local,DC=domain,DC=fr"))
            {
                Console.WriteLine("OK");
            }
            else
            {
                Console.WriteLine("FAIL");
            }

            u.FindByIdentity("s0500066", 0);
            Console.WriteLine("Move {0}", u.UserPrincipalName);
            if (u.Move_To_OU("OU=Logistique,OU=SG,OU=STAFF,OU=local,DC=domain,DC=fr"))
            {
                Console.WriteLine("OK");
            }
            else
            {
                Console.WriteLine("FAIL");
            }
            // UserFolders uF = new UserFolders();
            //uF.CreateDirectory(@"C:\CUBASE", "GG_STUDENT", 2);

            // Archives();           
            // STAFF("s1600109");
            // STAFF_Info();
            // STUDENT();
            // STAFF_CNAM("s1300140");

            // Structure_Xlsx_Config_Loader xlsL = new Structure_Xlsx_Config_Loader(@"\\localasrv\Code.xlsx");
            // xlsL.print_XML();
            // xlsL.print_OU();
            // xlsL.create_GROUPS();

            /*            
            LDap_Operator op = new LDap_Operator("domain.net", "OU=STUDENT,OU=local,DC=domain,DC=fr",            
                                                 "domain.net", "OU=STUDENT,OU=local,DC=domain,DC=fr");

            LDap_Operator op_Global = new LDap_Operator("domain.net", "OU=local,DC=domain,DC=fr",
                                                        "domain.net", "OU=local,DC=domain,DC=fr");

            */
            //op.User_Print_Groups(@"pauline.meyer@domain.net");
            // op.is_User_Member_of(@"pauline.meyer@domain.net", @"gg_adm");
            //System.Text.RegularExpressions.Regex GrpNameLike = new Regex(@"(?i)^GG_STUDENT+");
            //op.User_Remove_from_Group(@"pauline.meyer@domain.net", GrpNameLike);


            /*
            User_descriptor u;            
            foreach ( string userSam in op.Users())
            {                
                u = new User_descriptor(null, op.principalContext);
                if (u.FindByIdentity(userSam, 0))
                {
                    u.FindGroups();
                    string MainGroup = u.OU_path_to_group().Replace("", "GG_");
                    if (u.is_member_of(MainGroup))
                        Console.WriteLine("{0} is member of {1}", u.Name, MainGroup);
                    else
                    {
                        Console.WriteLine("{0} is not member of {1}!", u.Name, MainGroup);
                        System.Text.RegularExpressions.Regex GrpNameLike = new Regex(@"(?i)^GG_STUDENT+");
                        op.User_Remove_from_Group(u.Name, GrpNameLike);
                        int level = 2;
                        while (u.OU_path_to_group(level).Length > 0)
                        {
                            string Grp = u.OU_path_to_group(level++).Replace("", "GG_");
                            if (op_Global.Group_add_User(Grp, u.User.SamAccountName))
                                Console.WriteLine(" added to {0} -> ok", Grp);
                            else
                                Console.WriteLine(" added to {0} -> FAIL!!!", Grp);
                        }
                    }
                }
                u = null;                               
            }

            */

            /*
            User_descriptor u = new User_descriptor(null, op.principalContext);
            if (u.FindByIdentity("pauline.meyer@domain.net"))
            {
                string MainGroup = u.OU_path_to_group().Replace("", "GG_");
                u.FindGroups();
                if (u.is_member_of(MainGroup))
                    Console.WriteLine("{0} is member of {1}", u.Name, MainGroup);
                else
                {
                    Console.WriteLine("{0} is not member of {1}!", u.Name, MainGroup);
                    System.Text.RegularExpressions.Regex GrpNameLike = new Regex(@"(?i)^GG_STUDENT+");
                    op.User_Remove_from_Group(@"pauline.meyer@domain.net", GrpNameLike);
                    int level = 2;
                    while (u.OU_path_to_group(level).Length>0) {
                        string Grp = u.OU_path_to_group(level++).Replace("", "GG_");
                        if(op_Global.Group_add_User(Grp, u.User.SamAccountName))
                            Console.WriteLine(" added to {0} -> ok", Grp);
                        else
                            Console.WriteLine(" added to {0} -> FAIL!!!", Grp);
                    }
                }
            }
            else
                Console.WriteLine("User not found...");
            */
            Console.WriteLine("--------------- THE END ---------------");
            //Console.ReadLine();
        }
    }
}
