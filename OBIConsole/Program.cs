using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OBI;
using System.ServiceProcess;
using System.Threading;
using System.Configuration.Assemblies;
using System.ComponentModel;


namespace OBIServer
{

	[RunInstaller(true)]
	public class ObiInstaller : System.Configuration.Install.Installer
	{
		public ObiInstaller()
		{
			ServiceProcessInstaller process = new ServiceProcessInstaller();

			process.Account = ServiceAccount.LocalSystem;

			ServiceInstaller serviceAdmin = new ServiceInstaller();

			serviceAdmin.StartType = ServiceStartMode.Manual;
			serviceAdmin.ServiceName = "ObiService";
			serviceAdmin.DisplayName = "Obi Presence Integration Service";

			// Microsoft didn't add the ability to add a description for the services we are going to install
			// To work around this we'll have to add the information directly to the registry but I'll leave
			// this exercise for later.


			// now just add the installers that we created to our parents container, the documentation
			// states that there is not any order that you need to worry about here but I'll still
			// go ahead and add them in the order that makes sense.
			Installers.Add(process);
			Installers.Add(serviceAdmin);
		}

	}


	class Program : PresenceListener
	{
		protected Lync_Control m_Lync;
		protected Bcm_Control m_BCM;
		protected oLog m_log;

		protected OBIConfig m_Conf;

		protected bool m_bRestarting;

		protected Dictionary<string, ObiUser> m_UsersLync;
		protected Dictionary<string, ObiUser> m_UsersBcm;

		protected int m_change_token;
		protected string m_DAI;
		protected string m_DAIUser;
		protected string m_DAIUserPassword;
		protected string m_Urn;
		protected bool m_brun;
		protected bool m_bdebug;
		protected ManualResetEvent m_shutdownEvent;

		SyncDirection m_sync_direction;
		int m_use_bcm_directory;
		
		///////////////////////////////////////////////
		// PresenceListener methods

    public ObiUser getUserBCM(string sBcmId)
    {
      ObiUser u = null;
      lock (this)
      {
        if (m_UsersBcm.TryGetValue(sBcmId, out u) == false)
          u = null;
      }
      return u;
    }
    public ObiUser getUserLync(string sUri)
    {
      ObiUser u = null;
      lock (this)
      {
        if (m_UsersLync.TryGetValue(sUri, out u) == false)
          u = null;
      }
      return u;
    }
    public ObiUser getUser(string sId)
    {
      ObiUser u = getUserBCM(sId);
      if (u == null)
        u = getUserLync(sId);
      return u;
    }
    public void addUser(ObiUser u)
    {
      lock (this)
      {
        if (m_UsersBcm.ContainsKey(u.UserId))
          m_UsersBcm.Remove(u.UserId);
        m_UsersBcm[u.UserId] = u;
        if (m_UsersLync.ContainsKey(u.Uri))
          m_UsersLync.Remove(u.Uri);
        m_UsersLync[u.Uri] = u;
      }
    }
    public void removeUser(ObiUser u)
    {
      lock (this)
      {
        if (m_UsersBcm.ContainsKey(u.UserId))
          m_UsersBcm.Remove(u.UserId);
        if (m_UsersLync.ContainsKey(u.Uri))
          m_UsersLync.Remove(u.Uri);
      }
    }
    public bool isUserBCM(string sId)
    {
      lock (this)
      {
        return m_UsersBcm.ContainsKey(sId);
      }
    }
    public bool isUserLync(string sId)
    {
      lock (this)
      {
        return m_UsersLync.ContainsKey(sId);
      }
    }

		public int UserRefreshStart()
		{
			m_change_token++;
			return m_change_token;
		}

		public void UserRefreshDone()
		{
			lock (this)
			{ 
				// redo subscriptions & remove untouched users
				ObiUser u;
				List<ObiUser> lremove;
				lremove = new List<ObiUser>(256);
				// start presence
				foreach (System.Collections.Generic.KeyValuePair<string, ObiUser> kp in m_UsersBcm)
				{
					u = kp.Value;
					if (u.CheckToken != m_change_token)
					{
						m_log.Log("Removing user " + u.logId + " from presence in UserRefreshDone");
						lremove.Add(u);
						m_BCM.StopPresence(u);
						m_Lync.StopPresence(u);
					}
				}
				foreach (ObiUser ur in lremove)
				{
					try
					{
						ObiUser ul;
						m_UsersBcm.Remove(ur.UserId);
						if (m_UsersLync.TryGetValue(ur.Uri, out ul))
						{
 							// remove only if same instance (since extid changes may change binding)
							if (ur == ul)
							{
								m_log.Log("Removing lync user with id " + ur.Uri + " in UserRefreshDone");
								m_Lync.StopPresence(ur);
								m_UsersLync.Remove(ur.Uri);
							}
						}
					}
					catch (Exception er)
					{
						m_log.Log("Exception in clearing user " + ur.logId + " in UserRefreshDone:" + er);
					}
				}
				foreach (System.Collections.Generic.KeyValuePair<string, ObiUser> kp in m_UsersBcm)
				{
					u = kp.Value;
					m_log.Log("Subscribing user " + u.logId + " from presence in UserRefreshDone");
					m_BCM.StartPresence(u);
					m_Lync.StartPresence(u);
				}
			}
		}

		public SyncDirection getSyncDirection()
		{
			return m_sync_direction;
		}

		public void Lync_PresenceChanged(ObiUser usr, Lync_to_Bcm_Rule rule)
		{
			m_log.Log("Unimplemented - Lync_PresenceChanged:" + usr.logId + " - rule " + rule.rId);
		}
		public void Lync_PresenceChanged(ObiUser usr)
		{
			lock (this)
			{
				m_log.Log("Lync presence changed: " + usr.logId + " - " + usr.LS.BaseString);

				if (m_sync_direction == SyncDirection.BCM_2_Lync) // only BCM->Lync used
					return;

				if (usr.LS != null)
				{
					try
					{
						Lync_to_Bcm_Rule rule = null;
						if (usr.BS.Sticky)
						{
							m_log.Log("User " + usr.logId + " lync state change, sticky bcm state on, ignoring");
							return;
						}

						// TOCHECK: PROFILE GROUP BASED SETTING ???
						if (m_use_bcm_directory > 0)
						{
							rule = usr.ProfileGroup.findLync2BcmRule(usr.LS);
							if (rule != null)
							{
								m_BCM.SetPresenceDirectory(usr, rule.BcmOnline.Name);
							}
							return;
						}


						if (usr.ProfileGroup != null)
						{
							rule = usr.ProfileGroup.findLync2BcmRule(usr.LS);
							if (rule != null)
							{
								string sBcmPro = null;
								if (usr.BS.LoginState)
								{
									if (rule.BcmOnline != null)
									{
										m_log.Log("user " + usr.logId + " changing online profile to " + rule.BcmOnline.BcmID + "," + rule.BcmOnline.Name);
										sBcmPro = rule.BcmOnline.BcmID;
									}
								}
								else
								{
									if (rule.BcmOffline != null)
									{
										m_log.Log("user " + usr.logId + " changing offline profile to " + rule.BcmOffline.BcmID + "," + rule.BcmOffline.Name);
										sBcmPro = rule.BcmOffline.BcmID;
									}
								}
								if (Util.IsNullOrEmpty(sBcmPro))
								{
									m_log.Log("User " + usr.logId + " bcm rule empty - ignoring");
								}
								else
								{
									// check internal rules
									switch (sBcmPro)
									{
										case "_Talking": // future: PSI call state
											m_log.Log("User " + usr.logId + " bcm rule _Talking");
											break;
										case "_EndCall": // future: PSI call state
											m_log.Log("User " + usr.logId + " bcm rule _CallEnd");
											break;
										case "_PaperWork":
											m_log.Log("User " + usr.logId + " bcm rule _PaperWork");
											if (usr.BS.ServiceState != BcmServiceState.Paperwork)
												m_BCM.SetPresence(usr, BcmServiceState.Paperwork);
											else
												m_log.Log("User " + usr.logId + " bcm already paperwork");
											break;
										case "_Working":
											m_log.Log("User " + usr.logId + " bcm rule _Working");
											if (usr.BS.ServiceState != BcmServiceState.Working)
												m_BCM.SetPresence(usr, BcmServiceState.Working);
											else
												m_log.Log("User " + usr.logId + " bcm already working");
											break;
										case "_LateAdmin":
										case "_WrapUp":
											m_log.Log("User " + usr.logId + " bcm rule _WrapUp");
											if (usr.BS.ServiceState != BcmServiceState.WrapUp)
												m_BCM.SetPresence(usr, BcmServiceState.WrapUp);
											else
												m_log.Log("User " + usr.logId + " bcm already wrapup");
											break;
										case "_Previous":
											if (usr.prevBS != null)
											{
												if (usr.prevBS.Profile != null)
												{
													if (usr.prevBS.Profile.Type == BcmStateOption.Profile)
														m_BCM.SetPresence(usr, usr.prevBS.Profile.BcmID);
												}
											}
											break;
										default:
											if (usr.BS.Profile != null && usr.BS.Profile.BcmID == sBcmPro)
												m_log.Log("User " + usr.logId + " bcm profile already on, ignoring");
											else
												m_BCM.SetPresence(usr, sBcmPro);
											break;
									}
								}
							}
							else
								m_log.Log("User " + usr.logId + " no bcm rule - ignoring");
						}
						else
							m_log.Log("User " + usr.logId + " has no profilegroup, ignoring");
					}
					catch (Exception ex)
					{
						m_log.Log("Exception in Lync_PresenceChanged", ex);
					}
				}
			}
		}

		public void Lync_PresenceUpdated(ObiUser usr)
		{
			m_log.Log("Lync presence updated:" + usr.logId);
		}

		public void Bcm_PresenceChanged(ObiUser usr)
		{
			m_log.Log("BCM_PresenceChanged - no rule, unimplemented");
		}
		public void Bcm_PresenceChanged(ObiUser usr,Bcm_to_Lync_Rule rule)
		{
			lock (this)
			{
				try
				{
					m_log.Log("BCM presence changed: " + usr.logId + " - rule:" + rule.rId);

					if (usr.LS.Sticky)
					{
						m_log.Log("User " + usr.logId + " lync state sticky, ignoring bcm state");
						return;
					}
					if (usr.LS.LoginState)
					{
						if (rule.LyncOnline != null)
						{
							if (usr.LS.IsDifferent(rule.LyncOnline))
							{
								//m_Lync.SetPresence(usr, rule.LyncOnline, rule);
								m_Lync.SetPresence(usr, rule.LyncOnline, rule);
							}
							else
								m_log.Log("BCM change - no state change in Lync for " + usr.logId);
						}
						else
							m_log.Log("BCM change - no rule for online lync " + usr.logId);
					}
					else
					{
						if (rule.LyncOffline != null)
						{
							if (usr.LS.IsDifferent(rule.LyncOffline))
							{
								m_Lync.SetPresence(usr, rule.LyncOffline, rule);
							}
							else
								m_log.Log("BCM change - no state change in Lync");
						}
						else
							m_log.Log("BCM change - no rule for offline lync");
					}
				}
				catch (Exception e)
				{
					m_log.Log("Exception in Bcm_PresenceChanged: " + e);
				}
			}
		}
		public void Bcm_PresenceUpdated(ObiUser usr)
		{
			m_log.Log("Bcm presence updated:" + usr.logId);
		}

		public void Lync_UserPresenceStarted(ObiUser usr)
		{
			m_log.Log("Presence started:" + usr.logId);
		}
		public void Lync_UserPresenceTerminated(ObiUser usr)
		{
			m_log.Log("Presence terminated:" + usr.logId);
		}

		public void Bcm_UserPresenceStarted(ObiUser usr)
		{
		}

		public void Bcm_UserPresenceTerminated(ObiUser usr)
		{
		}

		public void RequestRestart(string reason)
		{
			if (reason == null) reason = "<unknown>";
			m_log.Log("Restart requested: " + reason);
			m_bRestarting = true;
		}

		public string GetParameter(string sName, string sDefValue)
		{
			return m_Conf.GetParameter(sName, sDefValue);
		}
		public int GetParameterI(string sName, int iDefValue)
		{
			return m_Conf.GetParameterI(sName, iDefValue);
		}

		public void srv_log(string s)
		{
			m_log.Log("SRV:" + s);
		}
		public void closelog()
		{
			m_log.close();
		}

		public int run(bool bDebug,ManualResetEvent evtService)
		{
			try
			{
				string stv;

				m_bdebug = bDebug;
				m_brun = true;
				m_bRestarting = false;
				m_shutdownEvent = evtService;
				m_change_token = 1;
				m_log = new oLog();

				m_UsersLync = new Dictionary<string, ObiUser>(8192);
				m_UsersBcm = new Dictionary<string, ObiUser>(8192);

				m_log.Log("Starting up...");

				m_Conf = new OBIConfig(m_log);
				if (m_Conf.readBaseConfig() == false)
				{
					m_log.Log("Error - cannot initialize, readBaseConfig fails");
					return -1;
				}
				m_Lync = new Lync_Control(this, m_log, m_Conf);
				m_BCM = new Bcm_Control(this, m_log, m_Conf);

				m_DAI = GetParameter("PSI_Url", "");
				m_DAIUser = GetParameter("PSI_User", "");
				m_DAIUserPassword = GetParameter("PSI_UserPassword", "");
				m_Urn = GetParameter("urn", "no urn");

				m_sync_direction = SyncDirection.Both;
				stv = GetParameter("sync_direction", "both");
				if (stv == "bcm")
					m_sync_direction = SyncDirection.Lync_2_BCM;
				if (stv == "lync")
					m_sync_direction = SyncDirection.BCM_2_Lync;

				try
				{
					m_use_bcm_directory = 0;
					stv = GetParameter("BCM_LyncStateAsEntry", "0");
					m_use_bcm_directory = Int32.Parse(stv);
				}
				catch(Exception epr)
				{
					m_use_bcm_directory = 0;
				}
				if (m_use_bcm_directory>0)
				{
					stv = GetParameter("BCM_DB_Directory", "none");
					if (stv == "none")
					{
						m_log.Log("BCM_LyncStateAsEntry specified but BCM_DB_Directory DSN source is missing");
						m_use_bcm_directory = 0;
					}
				}
				
				if (GetParameter("autoprovision", "true") == "true")
				{
					if (!m_Lync.Start(m_Urn, null))
					{
						m_log.Log("Lync startup failed");
						return -3;
					}
				}
				else
				{
					if (!m_Lync.Start(null, null))
					{
						m_log.Log("Lync startup failed");
						return -4;
					}
				}
				if (!m_BCM.Start(m_DAI, m_DAIUser, m_DAIUserPassword))
				{
					m_log.Log("Error - startup for BCM failed");
					return -2;
				}

				// start presence
				foreach (System.Collections.Generic.KeyValuePair<string, ObiUser> kp in m_UsersBcm)
				{
					m_BCM.StartPresence(kp.Value);
					m_Lync.StartPresence(kp.Value);
				}
				m_BCM.StartPresenceQueries();

				while (m_brun)
				{
					// health checks etc...
					TimeSpan delay = new TimeSpan(0, 0, 10);
					if (m_shutdownEvent.WaitOne(delay, true) == true)
					{
						m_log.Log("Shutdown requested - ending ....");
						break;
					}
					if (m_bRestarting)
					{
						m_log.Log("Restart requested - ending main loop");
						m_brun = false;
					}
				}
				m_Lync.ShutDown();
				if (m_bRestarting)
					return 1;
				return 0;
			}
			catch (Exception erun)
			{
				m_log.Log("Unknown exception in OBI::run, stopping" + erun);
			}
			return -5;
		}

		public void end()
		{
			try
			{
				m_Lync.ShutDown();
			}
			catch (Exception e)
			{
				m_log.Log("in ending phase", e);
			}
		}
	}


	public class Starter
	{
		static void Main(string[] args)
		{
			/*string localhost = System.Net.Dns.GetHostEntry("localhost").HostName;
			Console.WriteLine("Localhost=" + localhost);
			*/
			foreach (string s in args)
			{
				if (s.Contains("debug"))
				{
					Program p = new Program();
					ManualResetEvent hStop = new ManualResetEvent(false);
					p.run(true,hStop);
					return;
				}
			}
			ObiService os = new ObiService();
			ServiceBase.Run(os);
		}
	}

	class ObiService : System.ServiceProcess.ServiceBase
	{
		public ObiService()
		{
			// create a new timespan object with a default of 10 seconds delay.
			m_delay = new TimeSpan(0, 0, 0, 10, 0);
		}

		/// <summary>
		/// Set things in motion so your service can do its work.
		/// </summary>
		protected override void OnStart(string[] args)
		{
			// create our threadstart object to wrap our delegate method
			ThreadStart ts = new ThreadStart(this.ServiceMain);

			// create the manual reset event and set it to an initial state of unsignaled
			m_shutdownEvent = new ManualResetEvent(false);

			// create the worker thread
			m_thread = new Thread(ts);

			// go ahead and start the worker thread
			m_thread.Start();

			// call the base class so it has a chance to perform any work it needs to
			base.OnStart(args);
		}

		/// <summary>
		/// Stop this service.
		/// </summary>
		protected override void OnStop()
		{
			// signal the event to shutdown
			m_shutdownEvent.Set();

			// wait for the thread to stop giving it 10 seconds
			m_thread.Join(10000);

			// call the base class 
			base.OnStop();
		}

		/// <summary>
		/// 
		/// </summary>
		protected void ServiceMain()
		{
			int retV;
			while (true)
			{
				m_prog = new Program();

				try
				{
					retV = m_prog.run(false, m_shutdownEvent);
				}
				catch (Exception e)
				{
					m_prog.srv_log("OBI::Execute - startup failed with exception:" + e);
					retV = -10;
				}
				if (retV < 0)
				{ 
					// startup error - wait a while to start again
					m_prog.srv_log("OBI::Execute - startup failed, ending");
					base.Stop();
					break;
				}
				if (retV > 0)
				{
 					// request for restart - wait a while
					m_prog.srv_log("OBI::Execute - restart requested, waiting a while");
					Thread.Sleep(10000); // <conf
					m_prog.srv_log("OBI::Execute - requested restart begin");
					m_prog.closelog();
				}
				if (retV == 0)
				{
 					// normal shutdown
					Console.WriteLine("OBI::Execute - normal service shutdown");
					m_prog.end();
					break;
				}
			}
		}

		protected Program m_prog;
		protected Thread m_thread;
		protected ManualResetEvent m_shutdownEvent;
		protected TimeSpan m_delay;
	}


}
