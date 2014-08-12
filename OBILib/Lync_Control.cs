/**********************************************
* Application Endpoint Starter
*
*
*
*/
using System;
using Microsoft.Rtc.Collaboration;
using Microsoft.Rtc.Signaling;
using Microsoft.Rtc.Collaboration.Presence;
using System.Threading;
using System.Configuration;
using System.Collections.Generic;
using System.Net;
using System.Security.Cryptography.X509Certificates;

namespace OBI
{

	class internal_ObiUser
	{
		public UserEndpoint uep;

		public bool bSettingPresence;   // indicates that presence update towards lync has been started
		public bool bIsInitial;         // to ignore initial notification

		public bool bPhoneState;     // set for 'speaking' state handling
		public string sSpeakingTo;   // other end uri 
		public bool bTalking;        // talking/disconnected 
	}

	public class Lync_Control
	{

		// collaboration platform - initialized with automatic or manual provisioning, using server settings
		CollaborationPlatform m_CollPlat;
		// trusted application endpoint - initialized with automatic or manual provisioning - no specific settings
		ApplicationEndpoint m_AppEP;

		// interface for user state changes
		PresenceListener m_PL;

		// configuration values
		OBIConfig m_CFG;

		// log interface (console or file)
    oLog m_log;

    Dictionary<string, ObiUser> m_Subscribed;

		// wait events for startup/shutdown synch
		ManualResetEvent m_eStartupWait = new ManualResetEvent(false);
		ManualResetEvent m_eShutdownWait = new ManualResetEvent(false);

		bool m_bPlatformOK = false;
		bool m_bAppEpOK = false;

		//bool m_bNotifyInitial = false;

		// PresenceView (subscription handling)
		RemotePresenceView m_RP;

		bool m_bAutoProvision;      // autoprovisioning used
		string m_ServerAddress;     // lync server address (for user
		string m_AppName;           // app urn / app name

		List<UserEndpoint> m_UEps;
		System.Threading.Timer m_timer;

		Dictionary<string, BasicLyncState> m_bases;

		/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// public interface

    public Lync_Control(PresenceListener PL, oLog log, OBIConfig cfg)
		{
			m_log = log;
			m_PL = PL;
			m_CFG = cfg;
      m_Subscribed = new Dictionary<string, ObiUser>(8192);
      m_UEps = new List<UserEndpoint>(64);
			TimerCallback tcb;
			tcb = new TimerCallback(AppEP_Timer);
			m_timer = new Timer(tcb, null, 30000, 10000);

			m_bases = new Dictionary<string, BasicLyncState>(16);
			m_bases.Add("Offline", BasicLyncState.Offline);
			m_bases.Add("Online", BasicLyncState.Available);
			m_bases.Add("Busy", BasicLyncState.Busy);
			m_bases.Add("Away", BasicLyncState.Away);
			m_bases.Add("DoNotDisturb", BasicLyncState.DoNotDisturb);
			m_bases.Add("BeRightBack", BasicLyncState.BeRightBack);
			m_bases.Add("IdleOnline", BasicLyncState.IdleOnline);
			m_bases.Add("IdleBusy", BasicLyncState.IdleBusy);
			m_bases.Add("off-work", BasicLyncState.OffWork);

		}

		////////////////////////////////////////////////////////
		// startup the platform and application endpoint
		// - sAppId - use urn identifier for autoprovision, null or empty for manual provision (settings from app.config...)
		public bool Start(string sAppId, String sLyncServer)
		{
			try
			{
				if (sLyncServer != null && sLyncServer.Length > 0)
					m_ServerAddress = sLyncServer;
				else
					m_ServerAddress = m_CFG.GetParameter("serverfqdn", "no server"); //ConfigurationManager.AppSettings["ServerFqdn"];

				if (sAppId != null && sAppId.Length > 0)
				{
					m_AppName = sAppId;
					m_bAutoProvision = true;
				}
				else
				{
					m_bAutoProvision = false;
					m_AppName = m_CFG.GetParameter("applicationname", "obi"); //ConfigurationManager.AppSettings["applicationName"];
				}

				m_log.Log("AppEP_Control.Start {0},{1}", m_AppName, m_ServerAddress);

				if (m_bAutoProvision)
				{
					m_log.Log("Using automatic provisioning");
					AutoProvisionInit(m_AppName);
					WaitForStartup(60000);
					if (!m_bPlatformOK)
					{
						m_log.Log("Automatic provisioning for platform failed, startup returing failure");
						return false;
					}
					if (!m_bAppEpOK)
					{
						m_log.Log("Automatic provisioning for application failed, startup returing failure");
						return false;
					}
				}
				else
				{
					ServerPlatformSettings sp_set;

					// read settings for manual provisioning
					string localhost = System.Net.Dns.GetHostEntry("localhost").HostName;
					string scert;
					int port = int.Parse(m_CFG.GetParameter("listeningport", "36001"));
					string gruu = m_CFG.GetParameter("gruu", "no gruu");
					string lho = m_CFG.GetParameter("localhostname", null);
					if (lho != null && lho.Length > 0)
						localhost = lho;

					m_log.Log("Using manual provisioning:");
					m_log.Log("Settings app={0},localhost={1},port={2},gruu={3}", m_AppName, localhost, port, gruu);

					scert = m_CFG.GetParameter("LocalCertificate", null);
          if (scert == null)
          {
            sp_set = new ServerPlatformSettings(m_AppName, localhost, port, gruu, GetLocalCertificate());
          }
          else
          {
            if (scert.StartsWith("iss:"))
              sp_set = new ServerPlatformSettings(m_AppName, localhost, port, gruu, GetLocalCertificateByIssuer(scert.Substring(4)));
            else
              sp_set = new ServerPlatformSettings(m_AppName, localhost, port, gruu, GetLocalCertificateBySubject(scert));
          }

					m_CollPlat = new CollaborationPlatform(sp_set);
					m_CollPlat.BeginStartup(OnPlatformStartupCompleted, null);

					WaitForStartup(60000);
					if (!m_bPlatformOK)
					{
						m_log.Log("Manual provisioning for platform failed, startup returing failure");
						return false;
					}
					if (!m_bAppEpOK)
					{
						m_log.Log("Manual provisioning for application failed, startup returing failure");
						return false;
					}
				}
				m_log.Log("Startup OK");
				return true;
			}
			catch (Exception ex)
			{
				m_log.Log("Exception in AppEP_Control::Start", ex);
				return false;
			}
		}

		///////////////////////////////////////
		// Shutdown the application and platform
		//
		public void ShutDown()
		{
			try
			{
				m_log.Log("AppEP_Control:Shutdown");
				if (m_AppEP != null)
				{
					m_AppEP.BeginTerminate(OnApplicationEndpointTerminateCompleted, null);
				}
				else
					ShutDownPlatform();
				WaitForShutdown(30000);
			}
			catch (Exception ex)
			{
				m_log.Log("Shutdown failed", ex);
			}
		}

		///////////////////////////////////////////////////////
		// StartPresence 
		// Start presence interchange for the given user 
		// parameter obiuser - externally created user object
		//
		public bool StartPresence(ObiUser usr)
		{
			try
			{
				List<RemotePresentitySubscriptionTarget> lu;
				RemotePresentitySubscriptionTarget rt;

				if (IsUser(usr))
				{
					m_log.Log("Presence interchange for user {0},{1} has already been started.", usr.Uri, usr.UserId);
					return false;
				}
				
				lu = new List<RemotePresentitySubscriptionTarget>();
				rt = new RemotePresentitySubscriptionTarget(usr.Uri);

				rt.SubscriptionContext = null;
				lu.Add(rt);
				if (m_RP == null)
				{
					RemotePresenceViewSettings rpvset = new RemotePresenceViewSettings();
					rpvset.SubscriptionMode = RemotePresenceViewSubscriptionMode.Default;
					m_RP = new RemotePresenceView(m_AppEP, rpvset);
					m_RP.ApplicationContext = "OBIApp";
					m_RP.PresenceNotificationReceived += new EventHandler<RemotePresentitiesNotificationEventArgs>(m_RP_PresenceNotificationReceived);
					m_RP.SubscriptionStateChanged += new EventHandler<RemoteSubscriptionStateChangedEventArgs>(m_RP_SubscriptionStateChanged);
				}
				m_RP.StartSubscribingToPresentities(lu);
				m_log.Log("Presence subscription starting for user {0},{1}", usr.Uri, usr.UserId);
        m_Subscribed.Add(usr.Uri, usr);
				return true;
			}
			catch (Exception ex)
			{
				m_log.Log("StartPresence exception:", ex);
				return false;
			}
		}

		///////////////////////////////////////////////////////
		// User access (GetById, exists, etc..)
		public ObiUser GetUser(string sId)
		{
      ObiUser u = null;
      if (m_PL != null)
      {
        u = m_PL.getUser(sId);
        if (u == null)
        {
          string su = sId.ToLower();
          u = m_PL.getUser(su);
        }
      }
      return u;
		}
		public ObiUser GetUserByLyncUri(string sLyncUri)
		{
      string su = sLyncUri.ToLower();
      if (m_PL != null)
        return m_PL.getUserLync(su);
      return null;
		}
		public bool IsUser(ObiUser u)
		{
      if (m_Subscribed.ContainsKey(u.Uri))
        return true;
      return false;
		}
		public bool IsUser(string sId)
		{
      string sl;
      if (m_Subscribed.ContainsKey(sId))
        return true;
      sl = sId.ToLower();
      if (m_Subscribed.ContainsKey(sl))
        return true;
			return false;
		}

		///////////////////////////////////////////////////////////////
		// SetPresence
		// - Methods for setting the presence state for users
		// Establishes temporary userendpoint for the user and sets presence
		// using automatic publication. Userendpoint is terminated automatically
		// when presence publish is done
		// (tbd: cache for userendpoints? anyway setting presence once does not
		//  automatically mean that it would be set again soonishly...)
		//

		public bool SetPresence(ObiUser usr, LyncState inLs, Bcm_to_Lync_Rule rule)
		{
			LyncState ls = new LyncState();
			ls.Set(inLs);
			if (rule.rId.Contains("_Talking"))
			{
				if (usr.LS.Type == LyncStateOption.BasicState || usr.LS.Type == LyncStateOption.CustomState)
				{
					m_log.Log("Storing previous lync state " + usr.LS.BaseString + " for " + usr.logId);
					usr.prevLS.Set(usr.LS);
				}
				else
				{
					m_log.Log("Ignoring previous lync state " + usr.LS.BaseString + " for " + usr.logId);
				}
			}
			/*else
			{
				if (usr.LS.LoginState == false)
				{
					m_log.Log("Set offline presence, storing previous lync state " + usr.LS.BaseString + " for " + usr.UserId);
					usr.prevLS.Set(usr.LS);
				}
			}*/

			if (ls.Type == LyncStateOption.Internal)
			{
				if (ls.BaseString == "_Previous")
				{
					if (usr.prevLS != null)
					{
						m_log.Log("User " + usr.logId + " setting previous state on");
						ls.Set(usr.prevLS);
					}
					else
					{
						m_log.Log("User " + usr.logId + " has no previous state, cannot set Presence");
						return false;
					}
				}
			}
			if (ls.Type == LyncStateOption.BasicState)
			{
				m_log.Log("Setting basicstate " + ls.Id + " to user " + usr.logId);
				return SetPresence(usr, ls.BaseState);
			}
			if (ls.Type == LyncStateOption.CustomState)
			{
				List<LocalizedString> sL = new List<LocalizedString>(ls.LCIDString.Count);
				foreach (System.Collections.Generic.KeyValuePair<long, string> kv in ls.LCIDString)
				{
					LocalizedString locs = new LocalizedString(kv.Key, kv.Value);
					m_log.Log("Adding locstring " + kv.Key + "," + kv.Value + " to prescoll");
					sL.Add(locs);
				}
				m_log.Log("Setting customstate " + ls.Id + " to user " + usr.logId);
				return SetPresence(usr, sL, (int)ls.Availability);
			}
			return true;
		}
		
		// Set custom presence state
		// - susr  - user uri/bcmid
		// - state - any string
		// - avlvalue - availability value describing the activity 
		public bool SetPresence(string susr, string state, int avlvalue)
		{
			ObiUser u = GetUser(susr);
			return SetPresence(u, state, avlvalue);
		}
		// Set custom presence state
		// - state - any string
		// - avlvalue - availability value describing the activity 
		private bool SetPresence(ObiUser usr, List<LocalizedString> lTokens, int avlValue)
		{
			try
			{
				UserEndpoint up;
				UserEndpointSettings upset;

				upset = new UserEndpointSettings(usr.Uri, m_ServerAddress);
				upset.AutomaticPresencePublicationEnabled = true;
				upset.Presence.UserPresenceState = new PresenceState(PresenceStateType.UserState, avlValue, new PresenceActivity(lTokens));

				up = new UserEndpoint(m_CollPlat, upset);

				getInternal(usr).bSettingPresence = true;

				up.BeginEstablish(OnUserEndpointEstablished, up);



				return true; // meaning that started ok. todo: callback for failure/success ?
			}
			catch (Exception ex)
			{
				m_log.Log("SetPresence for userendpoint failed:", ex);
				return false;
			}
		}

		// Set custom presence state
		// - state - any string
		// - avlvalue - availability value describing the activity 
		public bool SetPresence(ObiUser usr, string state, int avlValue)
		{
			try
			{
				UserEndpoint up;
				UserEndpointSettings upset;

				upset = new UserEndpointSettings(usr.Uri, m_ServerAddress);
				upset.AutomaticPresencePublicationEnabled = true;
				upset.Presence.UserPresenceState = new PresenceState(PresenceStateType.UserState, avlValue, new PresenceActivity(new LocalizedString(state)));

				//upset.Presence.UserPresenceState = new PresenceState(PresenceStateType.UserState,avlValue,new PresenceActivity(


				up = new UserEndpoint(m_CollPlat, upset);

				getInternal(usr).bSettingPresence = true;

				up.BeginEstablish(OnUserEndpointEstablished, up);



				return true; // meaning that started ok. todo: callback for failure/success ?
			}
			catch (Exception ex)
			{
				m_log.Log("SetPresence for userendpoint failed:", ex);
				return false;
			}
		}
		// Set basic presence state
		// lyncState - basic lyncstate enumeration (available,busy,dnd..)
		public bool SetPresence(ObiUser usr, BasicLyncState lyncState)
		{
			try
			{
				UserEndpoint up;
				UserEndpointSettings upset;
				PresenceState ps;
				upset = new UserEndpointSettings(usr.Uri, m_ServerAddress);
				upset.AutomaticPresencePublicationEnabled = true;
				switch (lyncState)
				{
					case BasicLyncState.Available: ps = PresenceState.UserAvailable; break;
					case BasicLyncState.Busy: ps = PresenceState.UserBusy; break;
					case BasicLyncState.DoNotDisturb: ps = PresenceState.UserDoNotDisturb; break;
					case BasicLyncState.BeRightBack: ps = PresenceState.UserBeRightBack; break;
					case BasicLyncState.OffWork: ps = PresenceState.UserOffWork; break;
					case BasicLyncState.Away: ps = PresenceState.UserAway; break;

					default:
						//TODO: how to handle custom states....
						ps = PresenceState.UserAvailable;
						break;
				}
				//case 6: ps = PresenceState.User;break;

				upset.Presence.UserPresenceState = ps;
				up = new UserEndpoint(m_CollPlat, upset);
				getInternal(usr).bSettingPresence = true;
				up.BeginEstablish(OnUserEndpointEstablished, up);

				return true; // meaning that started ok. todo: callback for failure/success ?
			}
			catch (Exception ex)
			{
				m_log.Log("SetPresence for userendpoint failed:", ex);
				return false;
			}
		}
		// Set 'fake' In-a-Call state using Busy availability value and "In-A-Call" custom state string
		// - bTalking - set/clear
		public bool SetPresence(ObiUser usr, bool bTalking)
		{
			try
			{
				UserEndpoint up;
				UserEndpointSettings upset;

				upset = new UserEndpointSettings(usr.Uri, m_ServerAddress);
				upset.AutomaticPresencePublicationEnabled = true;
				upset.Presence.UserPresenceState = new PresenceState(PresenceStateType.UserState, 6500, new PresenceActivity(new LocalizedString("In-a-Call")));

				up = new UserEndpoint(m_CollPlat, upset);
				getInternal(usr).bSettingPresence = true;
				up.BeginEstablish(OnUserEndpointEstablished, up);

				return true; // meaning that started ok. todo: callback for failure/success ?
			}
			catch (Exception ex)
			{
				m_log.Log("SetPresence for userendpoint failed:", ex);
				return false;
			}
		}

		// Set real In-a-Call phone state using LocalPresence and PhoneState presence category
		public bool SetPresenceAsPhoneState(ObiUser usr, bool bTalking, string sNumber)
		{
			try
			{
				UserEndpoint up;
				UserEndpointSettings upset;
				internal_ObiUser iusr;
				upset = new UserEndpointSettings(usr.Uri, m_ServerAddress);
				upset.AutomaticPresencePublicationEnabled = true;

				up = new UserEndpoint(m_CollPlat, upset);
				up.StateChanged += UserEndpoint_OnStateChanged;

				iusr = getInternal(usr);
				iusr.bPhoneState = true;
				iusr.bTalking = bTalking;
				iusr.sSpeakingTo = "tel:" + sNumber;
				iusr.uep = up;

				up.BeginEstablish(OnUserEndpointEstablished, up);

				return true; // meaning that started ok. todo: callback for failure/success ?
			}
			catch (Exception ex)
			{
				m_log.Log("SetPresence for userendpoint failed:", ex);
				return false;
			}
		}


		///////////////////////////////////////////////////////////////
		// StopPresence
		// - remove presence subscription
		//
		public bool StopPresence(string sUri)
		{
			ObiUser u = GetUserByLyncUri(sUri);
			if (u == null) return false;
			return StopPresence(u);
		}
		public bool StopPresence(ObiUser usr)
		{
			try
			{
				List<string> lu;
				if (m_RP == null)
					return false;
				if (!IsUser(usr))
					return false;
				lu = new List<string>();
				lu.Add(usr.Uri);
				m_RP.StartUnsubscribingToPresentities(lu);
        m_Subscribed.Remove(usr.Uri);
				return true;
			}
			catch (Exception ex)
			{
				m_log.Log("Exception in StopPresence(" + usr.Uri + ")", ex);
				return false;
			}
		}


		////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// platform and application endpoint callbacks
		//


		////////////////////////////////////////////////////////////////////////
		// Platform startup done
		//
		private void OnPlatformStartupCompleted(IAsyncResult result)
		{
			try
			{
				m_CollPlat.EndStartup(result);
				m_log.Log("StartupCompleted: {0}", result);
				m_bPlatformOK = true;
				EstablishAppEndpoint();
			}
			catch (RealTimeException ex)
			{
				m_log.Log("Platform startup failed >" + ex);
				m_bPlatformOK = false;
				m_eStartupWait.Set();
			}
		}
		////////////////////////////////////////////
		// Application endpoint established
		//
		private void OnApplicationEndpointEstablishCompleted(IAsyncResult res)
		{
			try
			{
				m_log.Log("OnApplicationEndpointEstablishCompleted");
				m_AppEP.EndEstablish(res);
				m_bAppEpOK = true;
				m_log.Log("AppEP Contact uri:{0}", m_AppEP.OwnerUri);
				m_log.Log("AppEP Endpoint uri:{0}", m_AppEP.EndpointUri);
			}
			catch (RealTimeException ex)
			{
				m_log.Log("AppEndpoint establish failed >" + ex);
				m_bAppEpOK = false;
			}
			m_eStartupWait.Set();
		}

		////////////////////////////////////////////
		// Application endpoint terminated
		//
		private void OnApplicationEndpointTerminateCompleted(IAsyncResult res)
		{
			try
			{
				m_log.Log("OnApplicationEndpointTerminateCompleted");
				m_AppEP.EndTerminate(res);
			}
			catch (RealTimeException ex)
			{
				m_log.Log("AppEP Terminate failed:{0}", ex);
			}
			ShutDownPlatform();
		}

		////////////////////////////////////////////
		// Platform terminated
		//
		private void OnPlatformShutDownCompleted(IAsyncResult res)
		{
			try
			{
				m_log.Log("OnPlatformShutDownCompleted");
				m_CollPlat.EndShutdown(res);
			}
			catch (RealTimeException ex)
			{
				m_log.Log("PlatformShutdown failed:{0}", ex);
			}
			m_eShutdownWait.Set();
		}


		/////////////////////////////////////////////
		// Autoprovision completed
		//
		private void OnCollPlatAutoProvInitCompleted(IAsyncResult res)
		{
			try
			{
				m_log.Log("OnCollPlatAutoProvInitCompleted");
				m_CollPlat.EndStartup(res);
				m_bPlatformOK = true;
			}
			catch (Exception ex)
			{
				m_log.Log("CollaborationPlatform startup failed", ex);
				m_bPlatformOK = false;
			}
		}

		//////////////////////////////////////////////////////////////////////////
		// Application endpoint handler, begin establishing for new endpoint 
		//
		void AutoProvisionApplicationEndpoint(object sender, ApplicationEndpointSettingsDiscoveredEventArgs args)
		{
			CollaborationPlatform platform = sender as CollaborationPlatform;

			if (m_AppEP != null)
			{
				m_log.Log("Autoprovisioned application endpoint '{0}' already initialized. Ignoring request for '{1}'", m_AppEP.OwnerUri, args.ApplicationEndpointSettings.OwnerUri);
				return;
			}

			try
			{
				m_log.Log("AutoProvisionApplicationEndpoint '{0}' - initializing", args.ApplicationEndpointSettings.OwnerUri);

				//args.ApplicationEndpointSettings.UseRegistration = true;

				m_AppEP = new ApplicationEndpoint(platform, args.ApplicationEndpointSettings);
				m_AppEP.StateChanged += new EventHandler<LocalEndpointStateChangedEventArgs>(m_AppEP_StateChanged);
				m_AppEP.BeginEstablish(OnAppEpAutoInitCompleted, null);

			}
			catch (Exception ex)
			{
				m_log.Log("Errror initializing application endpoint", ex);
				m_bAppEpOK = false;
				m_eStartupWait.Set();
			}
		}

		///////////////////////////////////////////////
		// Application endpoint initialize complete
		//
		private void OnAppEpAutoInitCompleted(IAsyncResult res)
		{
			try
			{
				m_log.Log("OnAppEpAutoInitCOmpleted");
				m_AppEP.EndEstablish(res);
				m_bAppEpOK = true;
			}
			catch (Exception ex)
			{
				m_log.Log("Application endpoint autoprovisioning failed", ex);
				m_bAppEpOK = false;
			}
			m_eStartupWait.Set();
		}

		///////////////////////////////////////////////
		// Userendpoint initialized
		//
		private void OnUserEndpointEstablished(IAsyncResult res)
		{
			try
			{
				bool bAddToRemoval = true;
				UserEndpoint up = (UserEndpoint)res.AsyncState;
				m_log.Log("OnUserEndpointEstablished: " + up.EndpointUri + "," + up.OwnerUri);
				up.EndEstablish(res);

        ObiUser usr = GetUserByLyncUri(up.OwnerUri);
				if (usr != null && usr.lyncInternal != null)
				{
					internal_ObiUser iusr = getInternal(usr);
					if (iusr.bPhoneState) // needs presenceupdate
					{
						//SetPhoneState(usr);
						bAddToRemoval = false;
					}
				}
				if (bAddToRemoval)
				{
					lock (this)
					{
						m_UEps.Add(up);
					}
				}

			}
			catch (Exception ex)
			{
				m_log.Log("OnUserEndpointEstablished:", ex);
			}
		}

		//////////////////////////////////////////////////////
		// UserEndpoint_OnUpdatePresenceDone
		// - Presence published via beingpublish (during SetPresenceAsPhoneState)
		private void UserEndpoint_OnUpdatePresenceDone(IAsyncResult res)
		{
			try
			{
				UserEndpoint uep = res.AsyncState as UserEndpoint;
				m_log.Log("OnPresUpdateDone");
				uep.PresenceServices.EndUpdatePresenceState(res);
			}
			catch (Exception ex)
			{
				m_log.Log("Setting phone state failed:", ex);
			}
		}

		//////////////////////////////////////////////////////
		// UserEndpoint_TerminateDone
		// - Callback for successful clearing after SetPresence 
		private void UserEndpoint_TerminateDone(IAsyncResult res)
		{
			try
			{
				UserEndpoint uep = res.AsyncState as UserEndpoint;
				uep.EndTerminate(res);
				m_log.Log("Userendpoint " + uep.OwnerUri + " terminated");
			}
			catch (Exception ex)
			{
				m_log.Log("Exception in UserEndpoint_TerminateDone", ex);
			}
		}


		//////////////////////////////////////////////////////
		// timer callback
		// - periodically clears userendpoints
		void AppEP_Timer(object o)
		{
			lock (this)
			{
				foreach (UserEndpoint u in m_UEps)
				{
					m_log.Log("BeginTerminate for userendpoint:" + u.OwnerUri);
					u.BeginTerminate(UserEndpoint_TerminateDone, u);
				}
				m_UEps.Clear();
			}
		}


		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// Events

		//////////////////////////////////////////////////////////////////////
		// Handler for application endpoint state 
		private void m_AppEP_StateChanged(object sender, LocalEndpointStateChangedEventArgs e)
		{
			m_log.Log("AppEndpoint state changed from {0} to {1}", e.PreviousState, e.State);
		}


		/////////////////////////////////////////////
		// User presence notification received
		// - inform listener for either presence changed or presence update done
		//
		private void m_RP_PresenceNotificationReceived(object sender, RemotePresentitiesNotificationEventArgs e)
		{
			try
			{
				RemotePresenceView view = (RemotePresenceView)sender;
				ObiUser usr;
				if (m_PL == null)
				{
					m_log.Log("No listener - ignoring precencenotification");
					return;
				}
				foreach (RemotePresentityNotification notification in e.Notifications)
				{
					m_log.Log("PresentityNotification for User:" + notification.PresentityUri);
					usr = GetUserByLyncUri(notification.PresentityUri);
					/*  test - should there be some use for Lync 'Note' field ?? >>> 
					if (notification.PersonalNote != null && notification.PersonalNote.Message != null)
						m_log.Log("Note:" + notification.PersonalNote.Message);
					if (notification.AggregatedPresenceState != null)
					{
							m_log.Log("AVal:" + notification.AggregatedPresenceState.AvailabilityValue);
							m_log.Log("Ctgr:" + notification.AggregatedPresenceState.CategoryName);
							m_log.Log("Loca:" + notification.AggregatedPresenceState.EndpointLocation);
					}
					<< test & log */

					if (usr != null)
					{
						internal_ObiUser iu;
						LyncState lnew;
						bool bCheckCustom = true;

						iu = getInternal(usr);

						// store lync state to usr
						if (notification.AggregatedPresenceState == null)
						{
							m_log.Log("User " + usr.Uri + " aggregated presence null - ignoring notification");
							continue;
						}

						lnew = new LyncState();
						lnew.BaseString = notification.AggregatedPresenceState.Availability.ToString();
						lnew.Availability = notification.AggregatedPresenceState.AvailabilityValue;

						if (lnew.BaseString == "Offline")
							lnew.LoginState = false;
						else
							lnew.LoginState = true;
						
						m_log.Log("User:" + usr.Uri + " availability=" + lnew.BaseString + " , value=" + lnew.Availability);

						// TOCHECK: SHOULD BE PER PROFILEGROUP ???
						if (m_PL.getSyncDirection() == SyncDirection.BCM_2_Lync)
							continue;
						
						if (notification.AggregatedPresenceState.Activity != null)
						{
							if (notification.AggregatedPresenceState.Activity.ActivityToken != null)
							{
								m_log.Log("ActivityToken=" + notification.AggregatedPresenceState.Activity.ActivityToken);
								if (notification.AggregatedPresenceState.Activity.ActivityToken == "on-the-phone")
								{
									lnew.BaseString = "_Talking";
									lnew.CallState = true;
									lnew.Type = LyncStateOption.CallState;
									bCheckCustom = false;
								}
								if (notification.AggregatedPresenceState.Activity.ActivityToken == "off-work")
								{
									lnew.BaseString = "OffWork";
									lnew.Type = LyncStateOption.BasicState;
									bCheckCustom = false;
								}
							}
							if (bCheckCustom)
							{
								if (notification.AggregatedPresenceState.Activity.CustomTokens != null)
								{
									bool bset = false;
									if (lnew.CallState == false)
										lnew.Type = LyncStateOption.CustomState;
									foreach (LocalizedString ls in notification.AggregatedPresenceState.Activity.CustomTokens)
									{
										// TODO: get by user lcid
										m_log.Log("Localized in activities:" + ls.Value + " - " + ls.LCID);
										lnew.LCIDString.Add(ls.LCID, ls.Value);
										bset = true;
									}
									if (bset)
									{
										m_log.Log("Stored user " + usr.Uri + " state from aggregated activity token");
									}
								}
							}
						}
						if (lnew.Type == LyncStateOption.None) // not activity state
						{
							if (m_bases.TryGetValue(lnew.BaseString, out lnew.BaseState))
							{
								lnew.Type = LyncStateOption.BasicState;
							}
						}

						if (usr.lync_initial)
						{
							Lync_to_Bcm_Rule rule;
							m_log.Log("User " + usr.Uri + " initial lync availability received");
							usr.LS = lnew;
							usr.prevLS = new LyncState();
							usr.prevLS.Set(lnew);
							usr.lync_initial = false;
							rule = usr.ProfileGroup.findLync2BcmRule(usr.LS);
							if (rule != null)
							{
								if (rule.Sticky)
								{
									m_log.Log("User " + usr.Uri + " initial lync availability is sticky");
									usr.LS.Sticky = true;
								}
							}
						}
						else
						{
							if (usr.LS.IsDifferent(lnew))
							{
								if (usr.LS.CallState == true && lnew.CallState == false) //tocheck: need for 2 indications??
								{
									// need to get prevstate updated here....
									/*dbg*/ m_log.Log("_EndCall state for lyncuser " + usr.Uri + ", updating previous also");
									usr.prevLS.Set(lnew);
									
									lnew.BaseString = "_EndCall";
									lnew.CallState = false;
									lnew.Type = LyncStateOption.CallState;
								}
								if (iu.bSettingPresence)
								{
									m_log.Log("User " + usr.Uri + " was setting presence");
									usr.LS = lnew;/*LAST*/
									iu.bSettingPresence = false;
									m_PL.Lync_PresenceUpdated(usr);
								}
								else
								{
									usr.LS = lnew;
									m_PL.Lync_PresenceChanged(usr);
								}
							}
							else
							{
								if (iu.bSettingPresence)
								{
									m_log.Log("User " + usr.Uri + " was setting presence");
									iu.bSettingPresence = false;
									m_PL.Lync_PresenceUpdated(usr);
								}
								else
									m_log.Log("User " + usr.Uri + " lync presence notification, no state change");
							}
						}

					}
					else
						m_log.Log("Presencenotification for unknown user:" + notification.PresentityUri);
				}
			}
			catch (Exception ex)
			{
				m_log.Log("Exception in RemotePresenceView::PresenceNotificationReceived", ex);
			}
		}

		////////////////////////////////////
		// User subscription state changed
		// - inform listener for start/end states
		// 
		private void m_RP_SubscriptionStateChanged(object sender, RemoteSubscriptionStateChangedEventArgs e)
		{
      lock (this)
      {
        try
        {
          RemotePresenceView view = (RemotePresenceView)sender;
          ObiUser usr;

          m_log.Log("SubscriptionStateChanged");

          foreach (var kvp in e.SubscriptionStateChanges)
          {
            string sUri;
						sUri = kvp.Key.Uri.ToLower();
						m_log.Log("'{0}' received subscription state change for '{1}': '{2}' --> '{3}'", view.ApplicationContext, kvp.Key.Uri, kvp.Value.PreviousState, kvp.Value.State);

            usr = GetUserByLyncUri(sUri);
            if (kvp.Value.State == RemotePresentitySubscriptionState.Subscribed)
            {
              if (usr != null)
								m_PL.Lync_UserPresenceStarted(usr);
            }
            if (kvp.Value.State == RemotePresentitySubscriptionState.Terminated)
            {
              if (usr != null)
								m_PL.Lync_UserPresenceTerminated(usr);
							if (m_Subscribed.ContainsKey(sUri))
								m_Subscribed.Remove(sUri);
            }

          }
        }
        catch (Exception ex)
        {
          m_log.Log("Exception in RemotePresenceView::SubscriptionStateChanged", ex);
        }
      }
		}

		//////////////////////////////////////////////////////
		// uep_OnStateChanged
		// - Userendpoint statechange, used only during SetPresenceAsPhoneState if
		//   'real' phone state is being published
		private void UserEndpoint_OnStateChanged(object sender, LocalEndpointStateChangedEventArgs e)
		{
			try
			{
				UserEndpoint uep = sender as UserEndpoint;
				ObiUser usr = GetUser(uep.OwnerUri);
				internal_ObiUser iusr;

				m_log.Log("UserEndpoint state:" + e.State.ToString());
				if (usr != null && usr.lyncInternal != null)
				{
					if (e.State == LocalEndpointState.Established)
					{
						iusr = getInternal(usr);
						m_log.Log("UEP Established, setting phone state:" + usr.Uri);
						if (iusr.bPhoneState)
							SetPhoneState(usr);
					}
				}
			}
			catch (Exception ex)
			{
				m_log.Log("uep_OnStateChanged failed:", ex);
			}

		}



		////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// Internal logic (private methods called internally)

		/////////////////////////////////////////////////
		// Initialize autoprovisioning
		//
		private bool AutoProvisionInit(string applicationId)
		{
			try
			{
				ProvisionedApplicationPlatformSettings settings = new ProvisionedApplicationPlatformSettings("Obi App", applicationId);

				m_log.Log("Starting to initialize platform with application ID '{0}'", applicationId);

				m_CollPlat = new CollaborationPlatform(settings);

				// Register handler for initializing application endpoints
				m_CollPlat.RegisterForApplicationEndpointSettings(AutoProvisionApplicationEndpoint);

				m_CollPlat.BeginStartup(OnCollPlatAutoProvInitCompleted, null);
			}
			catch (Exception ex)
			{
				m_log.Log("CollaborationPlatform startup failed", ex);
				return false;
			}
			return true;
		}

		////////////////////////////////////////////////////////////////////
		// Intialize (establish) manually provisioned application endpoint
		//
		private void EstablishAppEndpoint()
		{
			try
			{
				string uri = m_CFG.GetParameter("contacturi", "no c-uri"); //ConfigurationManager.AppSettings["contactUri"];
				string srv = m_CFG.GetParameter("serverfqdn", "no server"); //ConfigurationManager.AppSettings["ServerFqdn"];
				int tlsport = 5061;

				ApplicationEndpointSettings appset = new ApplicationEndpointSettings(uri, srv, tlsport);
				appset.UseRegistration = true;
				appset.SetEndpointType(EndpointType.Application);

				m_AppEP = new ApplicationEndpoint(m_CollPlat, appset);
				m_AppEP.StateChanged += new EventHandler<LocalEndpointStateChangedEventArgs>(m_AppEP_StateChanged);

				m_log.Log("Begin establish with {0} {1} {2}", uri, srv, tlsport);
				m_AppEP.BeginEstablish(OnApplicationEndpointEstablishCompleted, null);
			}
			catch (Exception ex)
			{
				m_log.Log("AppEndpoint establish failed:", ex);
				m_bAppEpOK = false;
				m_eStartupWait.Set();
			}
		}

		////////////////////////////////////////////////////////////////////
		// ShutDownPlatform
		// Terminate platform 
		private void ShutDownPlatform()
		{
			try
			{
				m_log.Log("Shutting down collaboration platform");
				m_CollPlat.BeginShutdown(OnPlatformShutDownCompleted, null);
			}
			catch (Exception ex)
			{
				m_log.Log("Beginshutdown failed for platform", ex);
				m_eShutdownWait.Set();
			}
		}

		////////////////////////////////////////////////////////////////////
		// SetPhoneState
		// Used for setting 'real' phone device state, called after SetPresenceAsPhoneState
		// when endpoint state is established
		//
		private void SetPhoneState(ObiUser u)
		{
			try
			{
				UserEndpoint uep;
				internal_ObiUser iusr = getInternal(u);
				uep = iusr.uep;
				PresenceState ps = new PresenceState(6500, new PresenceActivity(new LocalizedString("On-the-phone")), PhoneCallType.Voip, iusr.sSpeakingTo);
				uep.LocalOwnerPresence.BeginPublishPresence(new PresenceCategory[] { ps }, UserEndpoint_OnUpdatePresenceDone, uep);
			}
			catch (Exception ex)
			{
				m_log.Log("Setting phone state failed:", ex);
			}
		}


		// Wait until the platform and application endpoint inform startup complete (other operations will fail before this)
		protected bool WaitForStartup()
		{
			return m_eStartupWait.WaitOne();
		}
		protected bool WaitForStartup(int ms)
		{
			return m_eStartupWait.WaitOne(ms);
		}
		// Wait until the platform and application endpoint inform shutdown complete (for graceful shutdown)
		protected bool WaitForShutdown()
		{
			return m_eShutdownWait.WaitOne();
		}
		protected bool WaitForShutdown(int ms)
		{
			return m_eShutdownWait.WaitOne(ms);
		}

		///////////////////////////////////////////////////////
		// get/create internal user object
		//
		internal_ObiUser getInternal(ObiUser u)
		{
			internal_ObiUser iu;
      lock(this)
      {
        if (u.lyncInternal == null)
        {
          iu = new internal_ObiUser();
          iu.bIsInitial = true;
          u.lyncInternal = iu;
        }
        else
          iu = u.lyncInternal as internal_ObiUser;
      }
			return iu;
		}


    //////////////////////////////////////////////////////////
    // Fetching certificate for the local machine (by hostname/subject/issuer)
    // 
    
    // plain hostname
    protected X509Certificate2 GetLocalCertificate()
    {
      try
      {
        X509Store store;
        X509Certificate2Collection certificates;
        string sHostName;

        sHostName = Dns.GetHostEntry("localhost").HostName.ToUpper();

        store = new X509Store(StoreLocation.LocalMachine);
        store.Open(OpenFlags.ReadOnly);
        certificates = store.Certificates;
        
        foreach (X509Certificate2 cert in certificates)
        {
          m_log.Log("Checking certificate: " + cert.SubjectName.Name);
          if (cert.SubjectName.Name.ToUpper().Contains(sHostName) && cert.HasPrivateKey)
          {
            m_log.Log("Using certificate: " + cert.SubjectName.Name + ", issuer=" + cert.Issuer);
            return cert;
          }
        }
        m_log.Log("No certificate can be matched for hostname:" + sHostName); 
      }
      catch (Exception e)
      {
        m_log.Log("Exception in GetLocalCertificate:" + e);
      }
      return null;
    }

    // plain subject
    protected X509Certificate2 GetLocalCertificateBySubject(string sDef)
    {
      try
      {
        X509Store store;
        X509Certificate2Collection certificates;
        
        store = new X509Store(StoreLocation.LocalMachine);
        store.Open(OpenFlags.ReadOnly);
        certificates = store.Certificates;
        foreach (X509Certificate2 cert in certificates)
        {
          m_log.Log("Checking certificate:" + cert.SubjectName.Name + " for subject:" + sDef);
          if (cert.SubjectName.Name.ToUpper().Contains(sDef.ToUpper()) && cert.HasPrivateKey)
          {
            m_log.Log("Using certificate: " + cert.SubjectName.Name + ", issuer=" + cert.Issuer);
            return cert;
          }
        }
        m_log.Log("No certificate can be matched for subject:" + sDef);
      }
      catch (Exception e)
      {
        m_log.Log("Exception in GetLocalCertificateBySubject:" + e);
      }

      return null;
    }

    // by hostname and issuer
    protected X509Certificate2 GetLocalCertificateByIssuer(string sDef)
    {
      try
      {
        X509Store store;
        X509Certificate2Collection certificates;
        string sHostName;
        
        sHostName = Dns.GetHostEntry("localhost").HostName.ToUpper();
        store = new X509Store(StoreLocation.LocalMachine);
        store.Open(OpenFlags.ReadOnly);
        certificates = store.Certificates;
        foreach (X509Certificate2 cert in certificates)
        {
          m_log.Log("Checking certificate: " + cert.SubjectName.Name + " for host:" + sHostName + " and issuer:" + cert.Issuer);
          if (cert.SubjectName.Name.ToUpper().Contains(sHostName) && cert.HasPrivateKey)
          {
            if (cert.Issuer != null && cert.Issuer.Contains(sDef))
            {
              m_log.Log("Using certificate: " + cert.SubjectName.Name + ", issuer=" + cert.Issuer);
              return cert;
            }
          }
        }
        m_log.Log("No certificate can be matched for host:" + sHostName + " and issuer:" + sDef);
      }
      catch (Exception e)
      {
        m_log.Log("Exception in GetLocalCertificateByIssuer:" + e);
      }
      return null;
    }

	}
}

