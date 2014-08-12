/**********************************************
* Application Endpoint Starter
*
*
*
*/
using System;
using System.Threading;
using System.Configuration;
using System.Collections.Generic;
using System.Net;
using System.IO;
using System.Xml;
using System.Linq;
using System.Text;
using System.Data.Odbc;
using Microsoft.Win32;


namespace OBI
{

	using OBILib.PSI;

	public class Bcm_Control
	{
		oLog m_log;
		PresenceListener m_PL;
		OBIConfig m_CFG;
		string m_sPSIUrl;
		string m_sPSIUser;
		string m_sPSIPass;
		PSI m_PSI;
		EventFlags m_EvtFlags;
		
		System.Net.CredentialCache m_CCache;

		Thread m_thread;

		string m_PSI_Session;

		Dictionary<string, ProfileGroup> m_Groups;

		Dictionary<string, ObiUser> m_Subscribed;

		public int m_RelevantChangeCountUsers;
		public int m_RelevantChangeCountGroups;
		public int m_RelevantChangeCountProfiles;

		public const int k_not_set=0;				// meaning not changed in notification
		public const int k_logout=10;
		public const int k_loggedin_cem=11;
		public const int k_loggedin_ext=12;

		public const int k_service_working=10;
		public const int k_service_paperwork=11;
		public const int k_service_wrapup=12;

		public const int k_callstate_idle=10;
		public const int k_callstate_talking = 11;

		protected int m_use_dir_entry = 0;

		Bcm_DB m_BcmDB;

    public Bcm_Control(PresenceListener PL, oLog log, OBIConfig cfg)
		{
			m_log = log;
			m_PL = PL;
			m_CFG = cfg;
			m_RelevantChangeCountUsers = 0;
			m_RelevantChangeCountGroups = 0;
			m_RelevantChangeCountProfiles = 0;
			m_PSI = null;
			m_PSI_Session = null;
			m_Subscribed = new Dictionary<string, ObiUser>(8192);
		}

		////////////////////////////////////////////////////////
		// Startup
		public bool Start(string sServerAddress,string sUserName,string sUserPassword)
		{
			m_PSI = new PSI();
			m_sPSIUrl = sServerAddress;
			m_PSI.Url = m_sPSIUrl;
			
			m_sPSIUser = sUserName;
			m_sPSIPass = sUserPassword;

			m_log.Log("BcmControl.Start PSIUrl=" + m_sPSIUrl);

			if (Util.IsNullOrEmpty(m_sPSIUser) || Util.IsNullOrEmpty(m_sPSIPass))
			{
				m_log.Log("Using empty credentials for PSI (no user/password set)");
			}
			else
			{
				m_log.Log("Using credentials for PSI " + m_sPSIUser + "," + m_sPSIPass);
				m_CCache = new System.Net.CredentialCache();
				m_CCache.Add(new System.Uri(m_sPSIUrl), "Basic", new System.Net.NetworkCredential(m_sPSIUser, m_sPSIPass));
				m_PSI.Credentials = m_CCache;
			}

			try
			{
				CreateSessionResult r;
				string stmp;

				r = m_PSI.CreateSession();

				m_PSI_Session = r.SessionID;
				m_log.Log("Created PSI session " + m_PSI_Session);


				if (GetProfileConfig() == false)
				{
					m_log.Log("Cannot get profile configuration, ending...");
					return false;
				}
				if (GetUserConfig()==false)
				{
					m_log.Log("Cannot get user configuration, ending...");
					return false;
				}

				m_EvtFlags = EventFlags.All;
				string stype = m_CFG.GetParameter("EventFlags","All");
				if (stype == "Call_and_Login")
					m_EvtFlags = EventFlags.Call_and_Login;
				if (stype == "Call")
					m_EvtFlags = EventFlags.Call;
				if (stype == "Call_and_Service")
					m_EvtFlags = EventFlags.Call_and_Service;
				if (stype == "Login")
					m_EvtFlags = EventFlags.Login;
				if (stype == "Login_and_Service")
					m_EvtFlags = EventFlags.Login_and_Service;

				stmp = m_CFG.GetParameter("BCM_LyncStateAsEntry", "0");
				if (stmp == "1")
				{
					stmp = m_CFG.GetParameter("BCM_DB_Directory", "none");
					if (stmp.Length > 0 && stmp != "none")
					{
						m_BcmDB = new Bcm_DB();
						m_BcmDB.init(m_CFG, m_log, stmp);
						m_use_dir_entry = 1;
					}
				}
			}
			catch (Exception x)
			{
				m_log.Log("Error in Bcm_Control start - cannot create session for PSI (" + x + ")");
				return false;
			}
			return true;
		}


		///////////////////////////////////////
		// Get configuration
		// (profiles from BCM, rules from cfg)
		protected bool GetProfileConfig()
		{ 
			try
			{
				PresenceProfileType[] aProf;

				// fetch the profiles & groups
				aProf = m_PSI.GetPresenceProfiles();

				foreach (PresenceProfileType pt in aProf)
				{
					m_log.Log("Profile guid=" + pt.ProfileID + " name=" + pt.ProfileName);
					BcmProfile bp = new BcmProfile();
					bp.BcmID = pt.ProfileID;
					bp.Name = pt.ProfileName;
					bp.Type = BcmStateOption.Profile;
					m_CFG.AddBCMProfile(bp);
				}
				if (m_CFG.readGroupConfig() == false)
				{
					m_log.Log("Profile configuration is faulty - cannot start");
					return false;
				}
				return true;
			}
			catch (Exception e)
			{
				m_log.Log("Exception in BCM_Control::FetchProfileConfig " + e);
			}
			return false;		
		}

		///////////////////////////////////////////////////
		// GetUserConfig
		// 
		protected bool GetUserConfig()
		{ 
			try
			{
				UserType[] aUsers;
				string[] aPresenceTokens;

				m_Groups = m_CFG.GetProfileGroups();
				aPresenceTokens = new string[1];
				foreach(System.Collections.Generic.KeyValuePair<string,ProfileGroup> pg in m_Groups)
				{
					//todo: possibly multiple tokens....
					aPresenceTokens[0] = pg.Value.m_id;
					aUsers = m_PSI.GetUsers(m_PSI_Session,aPresenceTokens);
					foreach (UserType u in aUsers)
					{
						ObiUser oU = new ObiUser();
						bool badd = true;
						oU.UserId = u.BcmID;
						oU.Uri = u.ExtID.ToLower();
						oU.PSToken = u.PSToken;
						oU.FirstName = u.FirstName;
						oU.SurName = u.SurName;
						oU.FName = u.FirstName + " " + u.SurName;
						oU.logId = "(" + oU.Uri + "," + oU.UserId + ")";
						oU.ProfileGroup = pg.Value;
						oU.BS.Profile = m_CFG.GetBcmUnkown();

						m_log.Log("Created user " + oU.UserId + " with lync id " + oU.Uri);

						if (m_PL.isUserBCM(oU.UserId))
						{
							m_log.Log("Error - user " + oU.logId + "(" + oU.FName + ") already in subscriptions (token=" + oU.PSToken + ")");
							badd = false;
						}
						if (m_PL.isUserLync(oU.Uri))
						{
							m_log.Log("Error - user ext " + oU.logId + "(" + oU.FName + ") already in subscriptions (token=" + oU.PSToken + ")");
							badd = false;
						}
						if (badd)
						{
              m_PL.addUser(oU);
						}
					}
				}
			}
			catch (Exception x)
			{
				m_log.Log("Error in Bcm_Control::GetUserConfig: " + x);
				return false;
			}
			return true;
		}

		///////////////////////////////////////
		// Refresh users
		//
		protected bool RefreshUsers(int checktoken)
		{
      try
      {
        UserType[] aUsers;
        string[] aPresenceTokens;
        ObiUser oU = null;

        m_Groups = m_CFG.GetProfileGroups();
        aPresenceTokens = new string[1];
        foreach (System.Collections.Generic.KeyValuePair<string, ProfileGroup> pg in m_Groups)
        {
          //todo: possibly multiple tokens....
          aPresenceTokens[0] = pg.Value.m_id;
					aUsers = m_PSI.GetUsers(m_PSI_Session,aPresenceTokens);
          foreach (UserType u in aUsers)
          {
            bool bnew = true;
            bool bmod = false;
            u.ExtID = u.ExtID.ToLower();
            if (m_PL.isUserBCM(u.BcmID))
            {
              bnew = false;
              oU = m_PL.getUserBCM(u.BcmID);
							oU.CheckToken = checktoken;
              if (oU.Uri != u.ExtID)
              {
                bmod = true;
                m_log.Log("User " + oU.logId + " external id changed to " + u.ExtID);
                if (m_PL.isUserLync(u.ExtID))
                {
                  m_log.Log("User " + oU.logId + " new uri " + u.ExtID + " had already uri bound to another user");
                }
              }
            }
            else
            {
              m_log.Log("User " + u.BcmID + "," + u.ExtID + " is new to OBI");
              bnew = true;
            }

            if (bnew) //TBC....
            {
              oU = new ObiUser();
              oU.UserId = u.BcmID;
              oU.Uri = u.ExtID;
              oU.PSToken = u.PSToken;
              oU.FirstName = u.FirstName;
              oU.SurName = u.SurName;
              oU.FName = u.FirstName + " " + u.SurName;
              oU.logId = "(" + oU.Uri + "," + oU.UserId + ")";
              oU.ProfileGroup = pg.Value;
							oU.CheckToken = checktoken;
              oU.BS.Profile = m_CFG.GetBcmUnkown();

              m_PL.addUser(oU);

              m_log.Log("Created user " + oU.UserId + " with lync id " + oU.Uri);
            }
            if (bmod)
            {
              if (oU != null)
              {
                m_log.Log("Modified user " + oU.UserId + " with lync id " + oU.Uri);
                m_PL.addUser(oU);
              }
            }

          }
        }
      }
      catch (Exception x)
      {
        m_log.Log("Error in Bcm_Control::GetUserConfig: " + x);
        return false;
      }
      return true;
		}

		///////////////////////////////////////
		// Refresh profile configuration
		//
		protected bool RefreshProfiles()
		{
			bool bret = true;
			try
			{
				PresenceProfileType[] aProf;

				m_log.Log("Refreshing BCM profiles");

				m_CFG.CreateBaseProfiles();

				// fetch the profiles & groups
				aProf = m_PSI.GetPresenceProfiles();

				foreach (PresenceProfileType pt in aProf)
				{
					m_log.Log("Profile guid=" + pt.ProfileID + " name=" + pt.ProfileName);
					BcmProfile bp = new BcmProfile();
					bp.BcmID = pt.ProfileID;
					bp.Name = pt.ProfileName;
					bp.Type = BcmStateOption.Profile;
					m_CFG.AddBCMProfile(bp);
				}
				m_CFG.readGroupConfig();
			}
			catch (Exception e)
			{
				m_log.Log("Exception in BCM_Control::FetchProfileConfig " + e);
				bret = false;
			}
			return bret;		
		}

		///////////////////////////////////////
		// Main thread
		//
		protected void PSI_MainThread()
		{
			int nRestartCount = 0;
			int nConfigCycle = 0;
			/*
			try
			{
				string[] atoks;
				atoks = new string[2];
				atoks[0] = "tiitaa";
				atoks[1] = "tuutuu";
				m_PSI.SubscribeConfigChanges(m_PSI_Session, atoks);

				atoks = new string[1];
				atoks[0] = "tuutuu";
				m_PSI.GetUsers(m_PSI_Session, atoks);

			}
			catch (Exception e)
			{
				m_log.Log("Exception while subscribing conf changes:" + e);
			}
			/*ENDTEST*/
			while (true)
			{
				Thread.Sleep(500);
				try
				{
					PSI_presencePoll();
				}
				catch(Exception e)
				{
					m_log.Log("Exception in PSI_MainThread: " + e);
					nRestartCount++;
				}
				if (nRestartCount > 3)
				{
					m_PL.RequestRestart("Error count exceeded in BCM presence thread");
          break;
				}
				nConfigCycle++;
				if (nConfigCycle > 2) // config?
				{
					try
					{
						PSI_changePoll();
						nConfigCycle = 0;
					}
					catch (Exception e)
					{
						m_log.Log("Exception in PSI_MainThread: " + e);
						nRestartCount++;
					}
					foreach (System.Collections.Generic.KeyValuePair<string, ProfileGroup> pg in m_Groups)
					{
						if (pg.Value.IsConfigChanged())
						{
							bool bgo = false;
							m_log.Log("Configuration for group " + pg.Key + " changed, rereading");
							// test with temporary object first
							try
							{
								ProfileGroup tmp = new ProfileGroup(pg.Key, m_log, m_CFG);
								tmp.CreateBaseRules();
								if (tmp.ReadConfig() == true)
									bgo = true;
							}
							catch (Exception e)
							{
								m_log.Log("Cannot prefetch changed configuration for group " + pg.Key + ", ignoring / exc:" + e);
							}
							if (bgo)
							{
								if (pg.Value.ReReadConfig() == false)
								{
									m_log.Log("Cannot re-read configuration for group " + pg.Key + " changed, rereading");
									// fatal ????
								}
								lock (m_PL)
								{
									RefreshProfiles();
								}
							}
 						}
					}
				}
			}
		}

		protected void PSI_changePoll()
		{
			ConfigurationChange[] aChanges;
			int cfgl;
			bool bRefresh = false;
			m_log.Log("Calling GetConfigChanges");
			aChanges = m_PSI.GetConfigurationChanges(m_PSI_Session);
			if (aChanges != null && aChanges.Length > 0)
			{
				foreach (ConfigurationChange c in aChanges)
				{
					m_log.Log("Confchange: otype=" + c.ObjectType + " change=" + c.ChangeId + " id=" + c.ObjectId);
					switch (c.ObjectType)
					{ 
						case "USER":
							if (c.ChangeId == "CREATED" || c.ChangeId == "DELETED")
								m_RelevantChangeCountUsers++;
							break;
						case "PSTOKEN":
							m_RelevantChangeCountGroups++;
							break;
						case "USER_PSID":
							m_RelevantChangeCountUsers++;
							break;
						case "PROFILE":
							m_RelevantChangeCountProfiles++;
							break;
					}
				}
			}
			cfgl = m_CFG.GetParameterI("RefreshLimitUsers", 50);
			if (m_RelevantChangeCountUsers > cfgl)
			{
				m_log.Log("Change count " + m_RelevantChangeCountUsers + " for users reached - refreshing settings from BCM");
				m_RelevantChangeCountUsers = 0;
				bRefresh = true;
			}
			cfgl = m_CFG.GetParameterI("RefreshLimitTokens", 50);
			if (m_RelevantChangeCountGroups > cfgl)
			{
				m_log.Log("Change count " + m_RelevantChangeCountGroups + " for tokens reached - refreshing settings from BCM");
				m_RelevantChangeCountGroups = 0;
				bRefresh = true;
			}
			cfgl = m_CFG.GetParameterI("RefreshLimitProfiles", 50);
			if (m_RelevantChangeCountProfiles > cfgl)
			{
				m_log.Log("Change count " + m_RelevantChangeCountProfiles + " for profiles reached - refreshing settings from BCM");
				m_RelevantChangeCountProfiles = 0;
				bRefresh = true;
			}
			if (bRefresh)
			{
				lock (m_PL)
				{
					int checktoken = m_PL.UserRefreshStart();
					RefreshProfiles();
					RefreshUsers(checktoken);
					m_PL.UserRefreshDone();
				}
			}
		}

		protected void PSI_presencePoll()
		{
			PresenceChangeNotification[] aChanges;
			m_log.Log("Calling GetPresenceChanges\n");
			aChanges = m_PSI.GetPresenceChanges(m_PSI_Session, 5000);
			if (aChanges != null && aChanges.Length > 0)
			{
				foreach (PresenceChangeNotification p in aChanges)
				{
					try
					{
						ObiUser u;
						u = m_PL.getUserBCM(p.UserId);
						if (u != null)
						{
							string rname = null;
							Bcm_to_Lync_Rule Rule = null;
							BcmState bnew = new BcmState();

							m_log.Log("BCM change:" + u.logId + " profile:" + p.PresenceProfile + " login:" + p.LoginState + " service:" + p.ServiceState + " call:" + p.CallState);

							if (m_PL.getSyncDirection() == SyncDirection.Lync_2_BCM)
								continue;

							if (p.PresenceProfile != null && p.PresenceProfile.Length > 0)
								bnew.Profile = m_CFG.GetBCMProfile(p.PresenceProfile);
							else
								bnew.Profile = u.BS.Profile;
							if (bnew.Profile == null)
								bnew.Profile = m_CFG.GetBcmUnkown();
							if (p.LoginState != UserLoginState.No_Change)
							{
								if (p.LoginState == UserLoginState.Logged_In)
								{
									rname = "_Online";
									bnew.LoginState = true;
								}
								else
								{
									rname = "_Offline";
									bnew.LoginState = false;
								}
							}
							else
								bnew.LoginState = u.BS.LoginState;
							if (p.ServiceState != UserServiceState.No_Change)
							{
								switch (p.ServiceState)
								{
									case UserServiceState.Working:
										rname = "_Working";
										bnew.ServiceState = BcmServiceState.Working; break;
									case UserServiceState.Paperwork:
										rname = "_PaperWork";
										bnew.ServiceState = BcmServiceState.Paperwork; break;
									case UserServiceState.Wrapup:
										rname = "_WrapUp";
										bnew.ServiceState = BcmServiceState.WrapUp; break;
									default:
										rname = "_Working";
										bnew.ServiceState = BcmServiceState.Working; break;
								}
							}
							else
							{
								bnew.ServiceState = u.BS.ServiceState;
							}
							if (p.CallState != UserCallState.No_Change)
							{
								if (p.CallState == UserCallState.Talking)
								{
									rname = "_Talking";
									bnew.CallState = true;
								}
								else
								{
									rname = "_EndCall";
									bnew.CallState = false;
								}
							}
							else
								bnew.CallState = u.BS.CallState;

							if (p.EndTime != DateTime.MaxValue && p.EndTime != DateTime.MinValue)
							{
								m_log.Log("Endtime in notification now " + p.EndTime);
							}

							if (u.bcm_initial == false)
							{
								m_log.Log("User " + u.logId + "(" + u.FName + ") presence changed");
								if (u.BS.IsDifferent(bnew))
								{
									// find rule first by change indication
									if (rname != null)
									{
										m_log.Log("Change indication rule check " + rname);
										Rule = u.ProfileGroup.findBcm2LyncRule(bnew.Profile.Name + rname);
										if (Rule == null)
											Rule = u.ProfileGroup.findBcm2LyncRule(rname);
									}
									if (Rule == null)
									{
										if (bnew.Profile != u.BS.Profile)
										{
											m_log.Log("Profile changed - rule by profile");
											Rule = u.ProfileGroup.findBcm2LyncRule(bnew.Profile.Name);
										}
										else
											m_log.Log("No profile change - no rule");
									}
									if (bnew.Profile == u.BS.Profile)
									{
										if (u.BS.Sticky)
										{
											m_log.Log("No profile change - sticky bcm on - keeping it");
											bnew.Sticky = true;
										}
 									}
									u.BS = bnew;
									if (Rule != null)
									{
										if (Rule.Sticky)
										{
											m_log.Log("Setting sticky BCM state ");
											u.BS.Sticky = true;
										}
										m_PL.Bcm_PresenceChanged(u, Rule);
									}
								}
							}
							else
							{
								m_log.Log("User " + u.logId + "(" + u.FName + ") initial presence received");
								u.bcm_initial = false;
								u.BS = bnew;
								u.prevBS = new BcmState();
								u.prevBS.Set(bnew);

								// check rule for stickiness
								if (bnew.Profile.Name != null)
								{
									Rule = u.ProfileGroup.findBcm2LyncRule(bnew.Profile.Name);
									if (Rule != null)
									{
										if (Rule.Sticky)
										{
											m_log.Log("User " + u.logId + " initial bcm availability is sticky");
											u.BS.Sticky = true;
										}
									}
								}
								if (m_Subscribed.ContainsKey(u.UserId) == false)
									m_Subscribed.Add(u.UserId, u);
							}
						}
						else
						{
							m_log.Log("Error - unknown user " + p.UserId + " in presencechanges");
						}
					}
					catch (Exception exp)
					{
						m_log.Log("Exceptoin in GetPresenceChanges: " + exp);
					}
				}
			}
			else
				m_log.Log("No presence changes");
		}


		///////////////////////////////////////
		// Shutdown
		//
		public void ShutDown()
		{
			try
			{
				if (m_PSI != null && m_PSI_Session != null)
					m_PSI.EndSession(m_PSI_Session);
			}
			catch (Exception x)
			{
				m_log.Log("Error in Bcm_Control shutdown (" + x + ")");
			}
		}

		///////////////////////////////////////////////////////
		// StartPresence 
		// Start presence interchange for the given user 
		// parameter obiuser - externally created user object
		//
		public bool StartPresence(ObiUser usr)
		{
			bool bret = true;
			m_log.Log("Adding user " + usr.logId + " for bcm presence");
			try
			{
				if (m_Subscribed.ContainsKey(usr.UserId) == false)
				{
					m_PSI.SubscribeUserPresence(m_PSI_Session, usr.UserId, m_EvtFlags);
					m_Subscribed.Add(usr.UserId, usr);
				}
				else
					m_log.Log("User " + usr.logId + " already subscribed");
			}
			catch (Exception x)
			{
				m_log.Log("StartPresence error:" + x);
				bret = false;
			}
			return bret;
		}


		///////////////////////////////////////////////////////
		// StartPresenceQueries
		// Starts presence change notification queries
		//
		public bool StartPresenceQueries()
		{
			ThreadStart ts = new ThreadStart(this.PSI_MainThread);

			// create the worker thread
			m_thread = new Thread(ts);

			// go ahead and start the worker thread
			m_thread.Start();
			
			return true;
		}


		///////////////////////////////////////////////////////////////
		// SetPresence
		// - Methods for setting the presence state for users
		public bool SetPresence(ObiUser usr, string sBcmProfile)
		{
			try
			{
				PresenceType pt = new PresenceType();
				pt.CallState = 0;
				pt.EndTime = DateTime.MaxValue;
				pt.ServiceState = UserServiceState.No_Change;
				pt.CallState = UserCallState.No_Change;
				/*test
				pt.EndTime = DateTime.Now;
				pt.EndTime.AddHours(1);
				pt.ServiceState = UserServiceState.Wrapup;
				test*/
				pt.PresenceProfile = sBcmProfile;
				if (usr.prevBS != null)
				{
					usr.prevBS.Set(usr.BS);
				}
				usr.BS.Profile = m_CFG.GetBCMProfile(sBcmProfile);
				m_log.Log("BCM::SetPresence:" + usr.logId + " profile=" + sBcmProfile);
				m_PSI.SetUserPresence(m_PSI_Session,usr.UserId, pt);
				m_PL.Bcm_PresenceUpdated(usr);
				return true;
			}
			catch(Exception e)
			{
				m_log.Log("Exception in BCM_Control::SetPresence: " + e);
			}
			return false;
		}

		///////////////////////////////////////////////////////////////
		// SetPresence
		// - Methods for setting the presence state for users
		public bool SetPresence(ObiUser usr, BcmServiceState nServiceState)
		{
			try
			{
				PresenceType pt = new PresenceType();
				pt.CallState = 0;
				pt.EndTime = DateTime.MaxValue;
				pt.ServiceState = (UserServiceState)nServiceState;
				pt.PresenceProfile = "";
				m_PSI.SetUserPresence(m_PSI_Session,usr.UserId, pt);
				m_PL.Bcm_PresenceUpdated(usr);
				return true;
			}
			catch(Exception e)
			{
				m_log.Log("Exception in BCM_Control::SetPresence: " + e);
			}
			return false;
		}


		///////////////////////////////////////////////////////////////
		// Set In-a-Call state
		public bool SetPresence(ObiUser usr, bool bTalking)
		{
			try
			{
				PresenceType pt = new PresenceType();
				if (bTalking)
					pt.CallState = UserCallState.Talking;
				else
					pt.CallState = UserCallState.Idle;
				pt.ServiceState = UserServiceState.No_Change;
				pt.PresenceProfile = "";
				m_PSI.SetUserPresence(m_PSI_Session,usr.UserId, pt);
				m_PL.Bcm_PresenceUpdated(usr);
			}
			catch(Exception e)
			{
				m_log.Log("Exception in BCM_Control::SetPresence: " + e);
			}
			return false;
		}


		///////////////////////////////////////////////////////////////
		// SetPresence
		// - Methods for setting the presence state for users
		public bool SetPresenceDirectory(ObiUser usr, string bcmValue)
		{
			bool bret = true;
			try
			{
				if (m_BcmDB != null)
				{
					if (m_CFG.GetParameter("BCM_Dir_AddStartTime", "0") == "1")
					{
						string stmp = bcmValue + " - " + DateTime.Now.ToString("HH:mm");
						m_BcmDB.AddUpdate(usr.UserId, stmp);
					}
					else
						m_BcmDB.AddUpdate(usr.UserId, bcmValue);
				}
				else
				{
					m_log.Log("BCM_Control::SetPresenceDirectory called - no DB configuration exists");
					bret = false;
				}
			}
			catch (Exception e)
			{
				m_log.Log("Exception in SetPresenceDirectory:" + e);
				bret = false;
			}
			return bret;
		}

		///////////////////////////////////////////////////////////////
		// StopPresence
		//
		public bool StopPresence(string sUri)
		{
			ObiUser u = m_PL.getUserLync(sUri);
			if (u == null) return false;
			return StopPresence(u);
		}
		public bool StopPresenceBcm(string sUserId)
		{
			ObiUser u = m_PL.getUserBCM(sUserId);
			if (u == null) return false;
			return StopPresence(u);
		}
		public bool StopPresence(ObiUser usr)
		{
			try
			{
				m_PSI.UnsubscribeUserPresence(m_PSI_Session, usr.UserId);
        m_Subscribed.Remove(usr.UserId);
			}
			catch (Exception e)
			{
				m_log.Log("Exception in BCM_Control::StopPresence: " + e);
			}
			return true;
		}

	}

	public class DBUpdate
	{
		public string sUser;
		public string sValue;
	}

	public class Bcm_DB
	{
		protected OBIConfig m_cfg;
    protected oLog m_log;
		protected string m_bcm_dir;
		protected bool m_brun;
		protected List<DBUpdate> m_DirUpdates;
		protected string m_DirMLGuid;
		Thread m_dth;

		protected Dictionary<string, string> m_User2Entry;

		public Bcm_DB()
		{
			m_DirMLGuid = null;
			m_DirUpdates = null;
			m_User2Entry = null;
		}

    public void init(OBIConfig cfg, oLog log, string sDB)
		{
			m_cfg = cfg;
			m_log = log;
			m_bcm_dir = sDB;
			m_DirUpdates = new List<DBUpdate>(256);
			m_User2Entry = new Dictionary<string, string>(2048);
			m_brun = true;
			ThreadStart ts = new ThreadStart(this.runDir);
			m_dth = new Thread(ts);
			m_dth.Start();
		}
		public void close()
		{
			m_brun = false;
			m_dth = null;
		}

		public void AddUpdate(string sUser, string sUpdate)
		{
			try
			{
				DBUpdate dbu = new DBUpdate();
				dbu.sUser = sUser;
				dbu.sValue = sUpdate;
				lock (this)
				{
					m_DirUpdates.Add(dbu);
				}
			}
			catch (Exception e)
			{
				m_log.Log("Exception in AddUpdate:" + e);
			}
		}

		protected int initEntryData(OdbcConnection od)
		{
			int r=0;
			try
			{
				OdbcDataReader odr;
				OdbcCommand oc;
				string sdbval;
				m_DirMLGuid = null;
				sdbval = m_cfg.GetParameter("BCM_Dir_Attribute", "Lync");
				string sq = "SELECT TOP 1 DirectoryMasterListGUID FROM DirectoryMasterListAttr WHERE Value='" + sdbval + "'";
				oc = new OdbcCommand(sq, od);
				odr = oc.ExecuteReader();
				while (odr.Read())
				{
					m_DirMLGuid = Convert.ToString(odr.GetValue(0));
				}
				odr.Close();
			}
			catch (Exception e)
			{
				m_log.Log("Exception in AddUpdate:" + e);
				r = -1;
			}
			return r;
		}


		protected string S2G(string sg)
		{
			try
			{
				string sr;
				if (sg.Contains('-'))
					return sg;
				//07E084C2-28CD-49B1-A5EB-97670DF89D86
				//07E084C228CD49B1A5EB97670DF89D86
				sr = sg.Substring(0, 8) + "-" + sg.Substring(8, 4) + "-" + sg.Substring(12, 4) + "-" + sg.Substring(16, 4) + "-" + sg.Substring(20);
				return sr;
			}
			catch (Exception e)
			{
				m_log.Log("Exception in S2G(" + sg + ") >" + e);
			}
			return sg;
		}


		protected string getUserDirEntry(OdbcConnection oD,string sUserGuid)
		{
			string sret = null;
			try
			{
				if (m_User2Entry.TryGetValue(sUserGuid, out sret) == false)
				{
					OdbcDataReader odr;
					OdbcCommand oc;
					string sdbval = null;
					string sproperguid = S2G(sUserGuid);
					string sq = "SELECT TOP 1 GUID FROM DirectoryEntry WHERE SourceGUID='" + sproperguid + "'";
					oc = new OdbcCommand(sq,oD);
					odr = oc.ExecuteReader();
					while (odr.Read())
					{
						sdbval = Convert.ToString(odr.GetValue(0));
					}
					odr.Close();
					if (sdbval != null)
					{
						m_User2Entry.Add(sUserGuid, sdbval);
						sret = sdbval;
						try
						{
							sq = "SELECT * FROM DirectoryEntryAttr WHERE DirectoryEntryGUID='" + sret + "' AND DirectoryMasterListGUID='" + m_DirMLGuid + "'";
							oc = new OdbcCommand(sq,oD);
							odr = oc.ExecuteReader();
							if (odr.HasRows)
							{
								odr.Close();
							}
							else
							{
								string sl = m_cfg.GetParameter("BCM_Dir_Language", "EN");
								odr.Close();
								sq = "INSERT INTO DirectoryEntryAttr(DirectoryEntryGUID,DirectoryMasterListGUID,MultiLingual,Language,TextValue,BinaryValue) VALUES ('" + sret + "','" + m_DirMLGuid + "',0,'" + sl + "','-',NULL)";
								oc = new OdbcCommand(sq,oD);
								oc.ExecuteNonQuery();
							}
						}
						catch (Exception er)
						{
							m_log.Log("Exception in insert entry:" + er);
						}
					}
				}
			}
			catch (Exception e)
			{
				m_log.Log("Exception in getUserDirEntry:" + e);
				sret = null;
			}
			return sret;
		}


		protected void runDir()
		{
			OdbcConnection oDir = null;
			string sConnDir = m_bcm_dir;
			while (m_brun)
			{
				while (oDir == null)
				{
					try
					{
						oDir = new OdbcConnection(sConnDir);
						oDir.Open();
					}
					catch (Exception eOD)
					{
						m_log.Log("Cannot open db connection to " + sConnDir + " -" + eOD);
						oDir = null;
					}
					if (oDir != null)
					{
						if (initEntryData(oDir) != 0)
							oDir = null;
					}
					if (oDir == null)
						Thread.Sleep(500);
				}
				while (m_brun)
				{
					int nu = 0;
					DBUpdate dbu;
					dbu = null;
					lock (this)
					{
						nu = m_DirUpdates.Count;
						if (nu > 0)
						{
							dbu = m_DirUpdates[0];
							m_DirUpdates.RemoveAt(0);
						}
					}
					if (dbu!=null)
					{
						int r = update_directory(oDir,dbu.sUser, dbu.sValue);
						if (r < 0)
						{
							try
							{
								if (oDir != null)
								{
									if (oDir.State == System.Data.ConnectionState.Open)
										oDir.Close();
								}
								oDir = null;
							}
							catch (Exception odcl)
							{
								m_log.Log("Exception during dirconnclose : " + odcl);
								oDir = null;
							}
							lock (this)
							{
								m_DirUpdates.Insert(0, dbu);
							}
						}
					}
				}
			}
			try
			{
				if (oDir != null)
				{
					oDir.Close();
					oDir = null;
				}
			}
			catch (Exception eoc)
			{
				m_log.Log("Exception in odbc close:" + eoc);
			}
		}

		protected int update_directory(OdbcConnection oDir,string sUser,string sValue)
		{
			int iret = 0;
			try
			{
				System.Data.Odbc.OdbcCommand oc;

				string sentry;
				string sq;

				sentry = getUserDirEntry(oDir,sUser);
				if (sentry != null)
				{
					sq = "UPDATE DirectoryEntryAttr SET TextValue='" + sValue + "' WHERE DirectoryMasterListGUID='" + m_DirMLGuid + "' AND DirectoryEntryGUID='" + sentry + "'";
					oc = new OdbcCommand(sq,oDir);
					oc.ExecuteNonQuery();
				}
			}
			catch(Exception e)
			{
				m_log.Log("Exception in update_directory > " + e);
				iret = -1;
			}
			return iret;
		}
	
	}




}