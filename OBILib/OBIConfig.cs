using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OBI;

namespace OBI
{
    public class OBIConfig
    {

			// for base config
			Dictionary<string, string> m_ConfValues;

			// for profile config groups
			Dictionary<string, ProfileGroup> m_Groups;

			// for BCM Profiles
			Dictionary<string, BcmProfile> m_BcmProfiles;

			BcmProfile m_BcmUnknown;

			oLog m_log;

			string m_path;

      public OBIConfig(oLog l)
			{
				m_log = l;
				m_ConfValues = new Dictionary<string, string>(32);
				m_Groups = new Dictionary<string, ProfileGroup>(32);
				try
				{
					m_path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
				}
				catch (Exception e)
				{
					Console.WriteLine("Cannot fetch the local path, defaulting to current directory\n ->" + e);
					m_path = ".";
				}
				CreateBaseProfiles();
			}

			public void CreateBaseProfiles()
			{
				m_BcmProfiles = new Dictionary<string, BcmProfile>(32);

				BcmProfile b;
				b = new BcmProfile();
				b.Type = BcmStateOption.Internal;
				b.Name = "_Talking";
				b.BcmID = "_Talking";
				m_BcmProfiles.Add(b.Name, b);

				b = new BcmProfile();
				b.Type = BcmStateOption.Internal;
				b.Name = "_EndCall";
				b.BcmID = "_EndCall";
				m_BcmProfiles.Add(b.Name, b);

				b = new BcmProfile();
				b.Type = BcmStateOption.Internal;
				b.Name = "_PaperWork";
				b.BcmID = "_PaperWork";
				m_BcmProfiles.Add(b.Name, b);

				b = new BcmProfile();
				b.Type = BcmStateOption.Internal;
				b.Name = "_WrapUp";
				b.BcmID = "_WrapUp";
				m_BcmProfiles.Add(b.Name, b);

				b = new BcmProfile();
				b.Type = BcmStateOption.Internal;
				b.Name = "_Offline";
				b.BcmID = "_Offline";
				m_BcmProfiles.Add(b.Name, b);

				b = new BcmProfile();
				b.Type = BcmStateOption.Internal;
				b.Name = "_Online";
				b.BcmID = "_Online";
				m_BcmProfiles.Add(b.Name, b);

				b = new BcmProfile();
				b.Type = BcmStateOption.Internal;
				b.Name = "_Previous";
				b.BcmID = "_Previous";
				m_BcmProfiles.Add(b.Name, b);

				b = new BcmProfile();
				b.Type = BcmStateOption.Internal;
				b.Name = "_unknown_";
				b.BcmID = "_unknown_";
				m_BcmUnknown = b;

			}

			public int GetProfileGroupCount()
			{
				return m_Groups.Count;
			}

			public Dictionary<string, ProfileGroup> GetProfileGroups()
			{
				return m_Groups;
			}

						
			public string GetParameter(string sName, string sDefValue)
			{
				if (m_ConfValues.ContainsKey(sName))
					return m_ConfValues[sName];
				return sDefValue;
			}

			public int GetParameterI(string sName, int iDefValue)
			{
				int ret = iDefValue;
				try
				{
					if (m_ConfValues.ContainsKey(sName))
						ret = Int32.Parse(m_ConfValues[sName]);
				}
				catch (Exception)
				{
					ret = iDefValue;
				}
				return ret;
			}

			public BcmProfile GetBCMProfile(String sId)
			{
				if (m_BcmProfiles.ContainsKey(sId))
					return m_BcmProfiles[sId];
				return null;
			}

			public BcmProfile GetBcmUnkown()
			{
				return m_BcmUnknown;
			}

			public void AddBCMProfile(BcmProfile bp)
			{
				if (m_BcmProfiles.ContainsKey(bp.BcmID))
					m_log.Log("Warning - bcmprofile " + bp.BcmID + " already in base config");
				else
					m_BcmProfiles.Add(bp.BcmID, bp);

				if (m_BcmProfiles.ContainsKey(bp.Name))
					m_log.Log("Warning - bcmprofile " + bp.Name + " already in base config");
				else
					m_BcmProfiles.Add(bp.Name, bp);

				// add also state-based/change profiles
				m_BcmProfiles.Add(bp.Name + "_Talking",bp);
				m_BcmProfiles.Add(bp.Name + "_Idle", bp);
				m_BcmProfiles.Add(bp.Name + "_PaperWork", bp);
				m_BcmProfiles.Add(bp.Name + "_Working", bp);
				m_BcmProfiles.Add(bp.Name + "_WrapUp", bp);
				m_BcmProfiles.Add(bp.Name + "_Offline", bp);
				m_BcmProfiles.Add(bp.Name + "_Online", bp);

			}

			public bool readBaseConfig()
			{
				XmlReader rd = null;
				try
				{
					string sfile;
					sfile = m_path + "\\OBIConf.xml";
					rd = XmlReader.Create(sfile);
					string sE, sA;
					ProfileGroup PG;
					bool bOK = false;
					
					while (rd.Read())
					{
						if (rd.IsStartElement())
						{
							sE = rd.Name;
							switch (sE)
							{
								case "environment":
									bOK = true;
									sA = rd.GetAttribute("urn");
									if (sA != null) m_ConfValues.Add("urn", sA);
									sA = rd.GetAttribute("serverfqdn");
									if (sA != null) m_ConfValues.Add("serverfqdn", sA);
									sA = rd.GetAttribute("autoprovision");
									if (sA != null) m_ConfValues.Add("autoprovision", sA);
									sA = rd.GetAttribute("applicationname");
									if (sA != null) m_ConfValues.Add("applicationname", sA);
									sA = rd.GetAttribute("fqdn");
									if (sA != null) m_ConfValues.Add("fqdn", sA);
									sA = rd.GetAttribute("gruu");
									if (sA != null) m_ConfValues.Add("gruu", sA);
									sA = rd.GetAttribute("listeningport");
									if (sA != null) m_ConfValues.Add("listeningport", sA);
									sA = rd.GetAttribute("contacturi");
									if (sA != null) m_ConfValues.Add("contacturi", sA);
									sA = rd.GetAttribute("dai");
									if (sA != null) m_ConfValues.Add("dai", sA);
									sA = rd.GetAttribute("dai_poll");
									if (sA != null) m_ConfValues.Add("dai_poll", sA);
									sA = rd.GetAttribute("PSI_Url");
									if (sA != null) m_ConfValues.Add("PSI_Url", sA);
									sA = rd.GetAttribute("PSI_User");
									if (sA != null) m_ConfValues.Add("PSI_User", sA);
									sA = rd.GetAttribute("PSI_UserPassword");
									if (sA != null) m_ConfValues.Add("PSI_UserPassword", sA);

									sA = rd.GetAttribute("LocalCertificate");
									if (string.IsNullOrEmpty(sA) == false)
										m_ConfValues.Add("LocalCertificate", sA);

									sA = rd.GetAttribute("localcertificate");
									if (string.IsNullOrEmpty(sA) == false)
										m_ConfValues.Add("LocalCertificate", sA);

									sA = rd.GetAttribute("BCM_LyncStateAsEntry");
									if (string.IsNullOrEmpty(sA) == false)
										m_ConfValues.Add("BCM_LyncStateAsEntry", sA);

									sA = rd.GetAttribute("BCM_DB_Directory");
									if (string.IsNullOrEmpty(sA) == false)
										m_ConfValues.Add("BCM_DB_Directory", sA);

									sA = rd.GetAttribute("BCM_Dir_Attribute");
									if (string.IsNullOrEmpty(sA) == false)
										m_ConfValues.Add("BCM_Dir_Attribute", sA);

									sA = rd.GetAttribute("BCM_Dir_Language");
									if (string.IsNullOrEmpty(sA) == false)
										m_ConfValues.Add("BCM_Dir_Language", sA);
									else
										m_ConfValues.Add("BCM_Dir_Language", "EN");

									sA = rd.GetAttribute("BCM_Dir_AddStartTime");
									if (string.IsNullOrEmpty(sA) == false)
										m_ConfValues.Add("BCM_Dir_AddStartTime", sA);

									sA = rd.GetAttribute("localhostname");
									if (string.IsNullOrEmpty(sA) == false)
										m_ConfValues.Add("localhostname", sA);

									sA = rd.GetAttribute("sync_direction");
									if (string.IsNullOrEmpty(sA) == false)
										m_ConfValues.Add("sync_direction", sA);

									sA = rd.GetAttribute("RefreshLimitUsers");
									if (sA != null)
									{
										m_ConfValues.Add("RefreshLimitUsers", sA);
									}
									sA = rd.GetAttribute("RefreshLimitTokens");
									if (sA != null)
									{
										m_ConfValues.Add("RefreshLimitTokens", sA);
									}
									sA = rd.GetAttribute("RefreshLimitProfiles");
									if (sA != null)
									{
										m_ConfValues.Add("RefreshLimitProfiles", sA);
									}
									sA = rd.GetAttribute("DailyRefresh");
									if (sA != null)
									{
										m_ConfValues.Add("DailyRefresh", sA);
									}
									sA = rd.GetAttribute("EventFlags");
									if (sA != null)
									{
										m_ConfValues.Add("EventFlags", sA);
									}
									break;
								case "group":
									sA = rd.GetAttribute("id");
									m_log.Log("ProfileGroup:" + sA);
									PG = new ProfileGroup(sA,m_log,this);
									m_Groups.Add(sA, PG);
									break;
								default:
									break;
							}
						}
					}
					rd.Close();
					if (bOK == false)
					{
						m_log.Log("No proper environment found in config - startup failed");
						return false;
					}
					if (m_Groups.Count == 0)
					{
						m_log.Log("No groups in config - startup failed");
						return false;
					}
					return true;
				}
				catch (Exception e)
				{
					m_log.Log("Exception in readConfig: ", e);
					try
					{
						if (rd != null)
							rd.Close();
					}
					catch (Exception er)
					{
						m_log.Log("Exception 2 in readConfig: ", er);
					}
				}
				return false;
			}


			public bool readGroupConfig()
			{
				try
				{
					ProfileGroup PG;
					foreach (System.Collections.Generic.KeyValuePair<string, ProfileGroup> it in m_Groups)
					{
						PG = it.Value;
						PG.CreateBaseRules();
						if (PG.ReadConfig() == false)
						{
							m_log.Log("Cannot read config for profilegroup " + PG.m_id);
							return false;
						}
					}
					return true;
				}
				catch (Exception e)
				{
					m_log.Log("Exception in readConfig: ", e);
				}
				return false;
			}

		}


		public class ProfileGroup
		{
			public string m_id;
			oLog m_log;
			OBIConfig m_cfg;

			public DateTime m_lastModTime;
			public string m_CfgFile;
			
			Dictionary<String, LyncState> m_LyncStates;
			Dictionary<String, BcmProfile> m_BcmStates;
			Dictionary<String, Lync_to_Bcm_Rule> m_Lync2Bcm;
			Dictionary<String, Bcm_to_Lync_Rule> m_Bcm2Lync;

			Dictionary<String, LyncState> m_LCID_CustomStates;

			string m_path;

			public ProfileGroup(String sId, oLog log,OBIConfig c)
			{
				m_id = sId;
				m_log = log;
				m_cfg = c;
				try
				{
					m_path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
				}
				catch (Exception e)
				{
					Console.WriteLine("Cannot fetch the local path, defaulting to current directory\n ->" + e);
					m_path = ".";
				}
			}

			public void CreateBaseRules()
			{
				m_LyncStates = new Dictionary<string, LyncState>(32);
				m_BcmStates = new Dictionary<string, BcmProfile>(32);
				m_Lync2Bcm = new Dictionary<string, Lync_to_Bcm_Rule>(32);
				m_Bcm2Lync = new Dictionary<string, Bcm_to_Lync_Rule>(32);
				m_LCID_CustomStates = new Dictionary<string, LyncState>(128);

				// fill in default lync states
				LyncState l;
				l = new LyncState();
				l.Type = LyncStateOption.BasicState;
				l.BaseState = BasicLyncState.Available;
				l.Id = "Online";
				l.BaseString = "Online";
				m_LyncStates.Add(l.Id,l);

				l = new LyncState();
				l.Type = LyncStateOption.BasicState;
				l.BaseState = BasicLyncState.Offline;
				l.Id = "Offline";
				l.BaseString = "Offline";
				m_LyncStates.Add(l.Id, l);

				l = new LyncState();
				l.Type = LyncStateOption.BasicState;
				l.BaseState = BasicLyncState.Busy;
				l.Id = "Busy";
				l.BaseString = "Busy";
				m_LyncStates.Add(l.Id,l);

				l = new LyncState();
				l.Type = LyncStateOption.BasicState;
				l.BaseState = BasicLyncState.BeRightBack;
				l.Id = "BeRightBack";
				l.BaseString = "BeRightBack";
				m_LyncStates.Add(l.Id,l);

				l = new LyncState();
				l.Type = LyncStateOption.BasicState;
				l.BaseState = BasicLyncState.Away;
				l.Id = "Away";
				l.BaseString = "Away";
				m_LyncStates.Add(l.Id,l);

				l = new LyncState();
				l.Type = LyncStateOption.BasicState;
				l.BaseState = BasicLyncState.DoNotDisturb;
				l.Id = "DoNotDisturb";
				l.BaseString = "DoNotDisturb";
				m_LyncStates.Add(l.Id,l);

				l = new LyncState();
				l.Type = LyncStateOption.BasicState;
				l.BaseState = BasicLyncState.OffWork; //tocheck: this seems to be localized action state in lync 2010
				l.Id = "OffWork";
				l.BaseString = "OffWork";
				m_LyncStates.Add(l.Id,l);

				l = new LyncState();
				l.Type = LyncStateOption.BasicState;
				l.BaseState = BasicLyncState.IdleOnline;
				l.Id = "IdleOnline";
				l.BaseString = "IdleOnline";
				m_LyncStates.Add(l.Id, l);

				l = new LyncState();
				l.Type = LyncStateOption.BasicState;
				l.BaseState = BasicLyncState.IdleBusy;
				l.Id = "IdleBusy";
				l.BaseString = "IdleBusy";
				m_LyncStates.Add(l.Id, l);

				l = new LyncState();
				l.Type = LyncStateOption.CallState;
				l.BaseState = BasicLyncState.Busy;
				l.Id = "_Talking";
				l.BaseString = "_Talking";
				m_LyncStates.Add(l.Id, l);

				l = new LyncState();
				l.Type = LyncStateOption.CallState;
				l.BaseState = BasicLyncState.Available;
				l.Id = "_EndCall";
				l.BaseString = "_EndCall";
				m_LyncStates.Add(l.Id, l);

				l = new LyncState();
				l.Type = LyncStateOption.Internal;
				l.BaseState = BasicLyncState.Available;
				l.Id = "_Previous";
				l.BaseString = "_Previous";
				
				m_LyncStates.Add(l.Id, l);
			}

			public bool ReReadConfig()
			{
				CreateBaseRules();
				return ReadConfig();
			}
			
			public bool ReadConfig()
			{
				XmlReader rd = null;
				try
				{
					string sE;
					m_CfgFile = m_path + "\\" + m_id + ".xml";
					rd = XmlReader.Create(m_CfgFile);
					try
					{
						m_lastModTime = System.IO.File.GetLastWriteTime(m_CfgFile);
					}
					catch (Exception ef)
					{
						m_log.Log("Cannot get last modification time for " + m_path + "\\" + m_id + ".xml" + ", defaulting to current" + ef);
						m_lastModTime = DateTime.Now;
					}
					while (rd.Read())
					{
						if (rd.IsStartElement())
						{
							sE = rd.Name;
							switch (sE)
							{
								case "LyncStates":
									readLyncStates(rd);
									break;
								case "BCM_to_Lync":
									readBcm2LyncRules(rd);
									break;
								case "Lync_to_BCM":
									readLync2BcmRules(rd);
									break;
								default:
									break;
							}
						}
					}
					rd.Close();
					return true;
				}
				catch (Exception e)
				{
					m_log.Log("Exception in readConfig: ", e);
					try
					{
						if (rd != null)
							rd.Close();
					}
					catch (Exception er)
					{
						m_log.Log("Exception 2 in readConfig: ", er);
					}
				}
				return false;
			}

			public bool IsConfigChanged()
			{
				try
				{
					DateTime dt = System.IO.File.GetLastWriteTime(m_CfgFile);
					if (dt.CompareTo(m_lastModTime)==0)
					{
						return false;
					}
					m_log.Log("IsConfigChanged " + m_CfgFile + " was " + m_lastModTime.ToString() + " is " + dt.ToString());
					return true;
				}
				catch (Exception e)
				{
					m_log.Log("Exception in IsConfigChanged: " + e);
				}
				return false;
			}

			private bool readLyncStates(XmlReader r)
			{
				LyncState ls;
				string sE, sA;
				ls = null;
				try
				{
					while (r.Read())
					{
						if (r.IsStartElement())
						{
							sE = r.Name;
							if (sE == "state")
							{
								ls = new LyncState();
								sA = r.GetAttribute("id");
								ls.Id = sA;
								ls.BaseString = sA;
								ls.Type = LyncStateOption.CustomState;
								sA = r.GetAttribute("availability");
								if (sA != null)
									ls.Availability = Int32.Parse(sA);
								m_log.Log("Adding lyncstate " + ls.Id);
								m_LyncStates.Add(ls.Id, ls);
							}
							if (sE == "lang")
							{
								String lcid, lname, lval;
								int lc;
								if (ls != null)
								{
									lname = r.GetAttribute("id");
									lcid = r.GetAttribute("lcid");
									lval = r.GetAttribute("text");
									ls.LangString.Add(lname, lval);
									lc = Int32.Parse(lcid);
									ls.LCIDString.Add(lc, lval);
									m_log.Log("Adding lyncstate LCID " + lcid + lval);
									m_LCID_CustomStates.Add(lcid + lval, ls);
									/* TBC..
									 * - This requires textual id's to be somewhat unique
									 * - Anyway only value to add would be availaibility
									 * - Formatted string search wouldn't work from lync (e.g. how to match: text="Meeting $EndTime$" )
									*/
								}
							}
							if (sE == "BCM_to_Lync")
							{
								readBcm2LyncRules(r);
								return true;
							}
							if (sE == "Lync_to_BCM")
							{
								readLync2BcmRules(r);
								return true;
							}
						}
					}
				}
				catch (Exception e)
				{
					m_log.Log("Exception while reading lyncstates " + e);
					return false;
				}
				return true;
			}

			private bool readLync2BcmRules(XmlReader r)
			{
				LyncState ls;
				Lync_to_Bcm_Rule R;
				string sE, sA;
				ls = null;
				R = null;
				try
				{
					while (r.Read())
					{
						if (r.IsStartElement())
						{
							sE = r.Name;
							if (sE == "state")
							{
								R = new Lync_to_Bcm_Rule();
								sA = r.GetAttribute("id");
								if (m_LyncStates.ContainsKey(sA))
								{
									ls = m_LyncStates[sA];
									R.rId = sA;
									R.LyncSt = ls;
									R.BcmOnline = null;
									R.BcmOnline = null;
									sA = r.GetAttribute("sticky");
									if (Util.IsNullOrEmpty(sA))
										R.Sticky = false;
									else
										if (sA == "1")
											R.Sticky = true;
									sA = r.GetAttribute("bcm_online");
									if (Util.IsNullOrEmpty(sA) == false)
										R.BcmOnline = m_cfg.GetBCMProfile(sA);
									sA = r.GetAttribute("bcm_offline");
									if (Util.IsNullOrEmpty(sA) == false)
										R.BcmOffline = m_cfg.GetBCMProfile(sA);
									if (m_Lync2Bcm.ContainsKey(R.rId)==false)
										m_Lync2Bcm.Add(R.rId, R);
									else
										m_log.Log("Error - rule " + R.rId + " already in lync2bcm states");
								}
								else
									m_log.Log("Error - invalid configuration element in Lync2BcmRules - no LyncState with name " + sA);
							}
							if (sE == "BCM_to_Lync")
							{
								readBcm2LyncRules(r);
								return true;
							}
							if (sE == "Lync_to_BCM")
							{
								readLync2BcmRules(r);
								return true;
							}
						}
					}
				}
				catch (Exception e)
				{
					m_log.Log("Exception while reading Lync2BCMRules " + e);
					return false;
				}
				return true;
			}

			private bool readBcm2LyncRules(XmlReader r)
			{
				Bcm_to_Lync_Rule R;
				BcmProfile bs;
				string sE, sA;
				bs = null;
				R = null;
				try
				{
					while (r.Read())
					{
						if (r.IsStartElement())
						{
							sE = r.Name;
							if (sE == "state")
							{
								R = new Bcm_to_Lync_Rule();
								sA = r.GetAttribute("id");
								bs = m_cfg.GetBCMProfile(sA);
								if (bs != null)
								{
									m_Bcm2Lync.Add(sA, R);
									R.rId = sA;
									R.BcmState = bs;
									R.LyncOffline = null;
									R.LyncOnline = null;
									sA = r.GetAttribute("sticky");
									if (Util.IsNullOrEmpty(sA))
										R.Sticky = false;
									else
										if (sA == "1")
											R.Sticky = true;
									sA = r.GetAttribute("lync_online");
									if (Util.IsNullOrEmpty(sA) == false && m_LyncStates.ContainsKey(sA))
										R.LyncOnline = m_LyncStates[sA];
									sA = r.GetAttribute("lync_offline");
									if (Util.IsNullOrEmpty(sA) == false && m_LyncStates.ContainsKey(sA))
										R.LyncOffline = m_LyncStates[sA];
								}
								else
									m_log.Log("Error - invalid configuration element in Bcm2LyncRules - no BCM profile with name " + sA);
							}
							if (sE == "BCM_to_Lync")
							{
								readBcm2LyncRules(r);
								return true;
							}
							if (sE == "Lync_to_BCM")
							{
								readLync2BcmRules(r);
								return true;
							}
						}
					}
				}
				catch (Exception e)
				{
					m_log.Log("Exception while reading bcm2lyncRules " + e);
					return false;
				}
				return true;
			}


			public Lync_to_Bcm_Rule findLync2BcmRule(LyncState LS)
			{
				LyncState confLS;
				Lync_to_Bcm_Rule rule;

				confLS = null;
				rule = null;
				
				// first match the state to ones configured for this group
				if (LS.Type == LyncStateOption.BasicState || LS.Type == LyncStateOption.CallState)
				{
					if (m_LyncStates.TryGetValue(LS.BaseString, out confLS) == false)
					{
						m_log.Log("Warning - BasicState " + LS.BaseString + " not part of group config");
					}
				}
				else // if (LS.Type == LyncStateOption.CustomState)
				{
					String srch;
					foreach(System.Collections.Generic.KeyValuePair<long,String> ls in LS.LCIDString)
					{
						srch = "" + ls.Key + ls.Value;
						if (m_LCID_CustomStates.TryGetValue(srch, out confLS))
						{
							m_log.Log("Found match for custom LCID string " + ls.Value + " on " + srch);
							break;
						}
					}
				}
				if (confLS != null)
				{
					if (m_Lync2Bcm.TryGetValue(confLS.Id, out rule))
						m_log.Log("Found rule " + rule.rId + " for state " + confLS.Id);
				}
				if (rule == null)
				{
					m_log.Log("Cannot match given live lyncstate to configuration set");
				}
				return rule;
			}


			public Bcm_to_Lync_Rule findBcm2LyncRule(string ruleName)
			{
 				Bcm_to_Lync_Rule rule = null;
				if (m_Bcm2Lync.TryGetValue(ruleName, out rule))
					return rule;
				m_log.Log("No matching rule for " + ruleName);
				return null;
			}

			public Bcm_to_Lync_Rule findBcm2LyncRule(BcmState BS)
			{
				Bcm_to_Lync_Rule rule = null;
				BcmProfile bp,confBS;
			
				// first match the state to ones configured for this group
				if (BS.Profile != null && BS.Profile.BcmID.Length > 0)
				{
					bp = BS.Profile;

					confBS = m_cfg.GetBCMProfile(bp.BcmID);
					if (confBS == null)
						confBS = m_cfg.GetBCMProfile(bp.Name);

					if (confBS != null)
					{
						if (m_Bcm2Lync.TryGetValue(confBS.Name, out rule))
							m_log.Log("Found matching rule for " + confBS.Name);
						else
							m_log.Log("No matching rule for " + confBS.Name);
					}
					else
						m_log.Log("No matching BCM profile state for " + bp.Name + " , " + bp.BcmID);
					//TBC: Default rule for BCM profile name -> lync custom state by lcid's
				}
				return rule;
			}


		}




}
