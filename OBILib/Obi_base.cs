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

namespace OBI
{
	public enum BasicLyncState
	{
		Available,
		Busy,
		DoNotDisturb,
		BeRightBack,
		OffWork,/*<<custom*/
		Away,
		Offline,
		IdleOnline,
		IdleBusy,
		None
	}

	public enum LyncStateOption
	{
		BasicState,
		CustomState,
		CallState,
		Internal,
		None
	}
	public enum BcmStateOption
	{
		Profile,
		Internal,
		None
	}
	public enum BcmServiceState
	{
 		Working = 10,
		Paperwork = 11,
		WrapUp = 12,
		None
	}

	public enum SyncDirection
	{ 
		Both,
		Lync_2_BCM,
		BCM_2_Lync
	}

	public class LyncState
	{
		public LyncState()
		{
			Id = "";
			Type = LyncStateOption.None;
			Availability = 0;
			BaseState = BasicLyncState.None;
			LangString = new Dictionary<string, string>(8);
			LCIDString = new Dictionary<long, string>(8);
			CallState = LoginState = Sticky = false;
		}
		public string Id;
		public LyncStateOption Type;
		public long Availability;
		public BasicLyncState BaseState;
		public string BaseString;
		public Dictionary<string, string> LangString;
		public Dictionary<long, string> LCIDString;
		public bool Sticky;
		public bool LoginState;
		public bool CallState;

		public bool IsDifferent(LyncState lnew)
		{
			if (Id != lnew.Id)
				return true;
			if (Type != lnew.Type)
				return true;
			if (BaseState != lnew.BaseState)
				return true;
			if (Availability != lnew.Availability)
				return true;
			if (BaseString != lnew.BaseString)
				return true;
			if (LoginState != lnew.LoginState)
				return true;
			return false;
		}

		public void Set(LyncState lNew)
		{
			Id = lNew.Id;
			Type = lNew.Type;
			Availability = lNew.Availability;
			BaseState = lNew.BaseState;
			BaseString = lNew.BaseString;
			LCIDString.Clear();
			foreach(System.Collections.Generic.KeyValuePair<long,string> ls in lNew.LCIDString)
			{
				LCIDString.Add(ls.Key, ls.Value);
			}
			Sticky = lNew.Sticky;
			LoginState = lNew.LoginState;
			CallState = lNew.CallState;
		}

	}

	public class BcmProfile
	{
		public BcmProfile()
		{
			BcmID = "";
			Name = "";
			Type = BcmStateOption.None;
		}
		public string BcmID;		//guid
		public string Name;			//config name
		public BcmStateOption Type;
	}

	public class BcmState
	{
		public BcmState()
		{
			Profile = new BcmProfile();
			ServiceState = BcmServiceState.None;
			CallState = LoginState = Sticky = false;
		}
		public BcmProfile Profile;
		public BcmServiceState ServiceState;
		public bool CallState;
		public bool LoginState;

		public bool Sticky;

		public bool IsDifferent(BcmState bnew)
		{
			if (Profile != null)
			{
				if (bnew.Profile != null)
				{
					if (Profile.BcmID != bnew.Profile.BcmID)
						return true;
				}
			}
			else
				if (bnew.Profile != null)
					return true;
			if (ServiceState != bnew.ServiceState)
				return true;
			if (CallState != bnew.CallState)
				return true;
			if (LoginState != bnew.LoginState)
				return true;
			return false;
		}

		public void Set(BcmState bNew)
		{
			Profile.BcmID = bNew.Profile.BcmID;
			Profile.Name = bNew.Profile.Name;
			Profile.Type = bNew.Profile.Type;
			ServiceState = bNew.ServiceState;
			CallState = bNew.CallState;
			LoginState = bNew.LoginState;
			Sticky = bNew.Sticky;
		}

	}


	public class Bcm_to_Lync_Rule
	{
		public Bcm_to_Lync_Rule()
		{
 			rId = null;
			BcmState = null;
			LyncOnline = null;
			LyncOffline = null;
			Sticky = false;
		}
		public string rId;
		public BcmProfile BcmState;
		public LyncState LyncOnline;
		public LyncState LyncOffline;
		public bool Sticky;
	}

	public class Lync_to_Bcm_Rule
	{
		public Lync_to_Bcm_Rule()
		{
			rId = null;
			LyncSt = null;
			BcmOnline = null;
			BcmOffline = null;
			Sticky = false;
		}

		public string rId;
		public LyncState LyncSt;
		public BcmProfile BcmOnline;
		public BcmProfile BcmOffline;
		public bool Sticky;
	}

	public static class Util
	{
		public static bool IsNullOrEmpty(string s)
		{
			if (s == null) return true;
			if (s.Length == 0) return true;
			return false;
		}
	}

	/*public interface ObiConfig
	{

		string GetParameter(string sName, string sDefValue);
		int GetParameterI(string sName, int iDefValue);
	}*/

	public interface PresenceListener
	{
		// presence changed in lync side
		void Lync_PresenceChanged(ObiUser usr,Lync_to_Bcm_Rule rule);
		void Lync_PresenceChanged(ObiUser usr);
		// presence update to lync done 
		void Lync_PresenceUpdated(ObiUser usr);

		// presence changed in bcm side
		void Bcm_PresenceChanged(ObiUser usr,Bcm_to_Lync_Rule rule);
		void Bcm_PresenceChanged(ObiUser usr);
		// presence update to bcm done
		void Bcm_PresenceUpdated(ObiUser usr);

		// presence exchange for the user has started
		void Lync_UserPresenceStarted(ObiUser usr);
		// presence exchange for the user has ended (either manually or due to invalid user)
		void Lync_UserPresenceTerminated(ObiUser usr);

		// presence exchange for the user has started
		void Bcm_UserPresenceStarted(ObiUser usr);
		// presence exchange for the user has ended  (either manually or due to invalid user)
		void Bcm_UserPresenceTerminated(ObiUser usr);

    ObiUser getUserBCM(string sBcmId);
    ObiUser getUserLync(string sUri);
    ObiUser getUser(string sId);
    bool isUserBCM(string sId);
    bool isUserLync(string sId);
    void addUser(ObiUser u);
    void removeUser(ObiUser u);

		int UserRefreshStart();
		void UserRefreshDone();

		void RequestRestart(string reason);

		SyncDirection getSyncDirection();
	}

	public class ObiUser
	{
		public ObiUser()
		{
			Uri = UserId = PSToken = FirstName = SurName = FName = "";
			bcm_initial = lync_initial = true;
			LS = new LyncState();
			BS = new BcmState();
			prevLS = null;
			prevBS = null;
			ProfileGroup = null;
			CheckToken = 0;
		}

		public string Uri;				// Lync uri
		public string UserId;     // Bcm  userid

		public string PSToken;		// group id in bcm
		public string FirstName;
		public string SurName;
		public string FName; // first+sur for convenience

		public bool bcm_initial;	// first notification coming...
		public bool lync_initial; // first notification coming...

		public LyncState LS;			// current lyncstate
		public BcmState  BS;			// current bcmstate

		public DateTime BcmEndTime;

		public LyncState prevLS;	// previous state (for set-reset)
		public BcmState prevBS;

		public ProfileGroup ProfileGroup;

		public String logId;

		public int CheckToken;

		public Object lyncInternal;    // stores controller private data
		public Object bcmInternal;     // stores controller private data

	}



}
