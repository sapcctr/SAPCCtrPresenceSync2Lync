using System;

namespace OBI
{

       
    public class oLog 
    {
        System.IO.StreamWriter m_f;
        String m_fn;
				String m_path;
        int m_lasthour;

        public oLog()
        {
            m_lasthour = 25;
            m_f = null;
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

				public void close()
				{
          lock (this)
          {
            try
            {
              if (m_f != null)
                m_f.Close();
              m_f = null;
            }
            catch (Exception e)
            {
              Console.WriteLine("Exc in log close:" + e);
            }
          }
				}

        public void _tm()
        {
            DateTime t = DateTime.Now;
            if (m_f != null)
							m_f.Write(t.ToString("HH:mm:ss.fff "));
                //m_f.Write("" + t.Hour + ":" + t.Minute + ":" + t.Second + "." + t.Millisecond + " ");
        }

        public void _checkFile()
        {
            DateTime dt = DateTime.Now;
            if (m_lasthour != dt.Hour)
            {
                if (m_f != null) 
                    m_f.Close();
                m_f = null;
            }
            if (m_f == null)
            {
								string sd = dt.ToString("yyyyMMdd_HH");	
                
								m_fn = m_path + "\\Logs\\OBI_" + sd + ".log"; 
                m_f = new System.IO.StreamWriter(m_fn, true);
                m_lasthour = dt.Hour;
            }
            if (m_f != null)
                m_f.Flush();
        }
        
        public void Log(String msg)
        {
          lock (this)
          {
            try
            {
              _checkFile();
              _tm();
              m_f.WriteLine(msg);
              Console.WriteLine(msg);
            }
            catch (Exception e)
            {
              m_f = null;
              Console.WriteLine("Exc in log:" + e);
            }
          }
        }
        public void Log(String msg, Exception ex)
        {
          lock (this)
          {
            try
            {
              _checkFile();
              _tm();
							m_f.WriteLine(msg + " >" + ex);
							Console.WriteLine(msg + " >" + ex);
							//m_f.WriteLine(msg + " >" + ex.Message);
              //Console.WriteLine(msg + " >" + ex.Message);
            }
            catch (Exception e)
            {
              m_f = null;
              Console.WriteLine("Exc in log:" + e);
            }
          }
        }

        public void Log(String msg, params object[] arg)
        {
          lock (this)
          {
            try
            {
              _checkFile();
              _tm();
              m_f.WriteLine(msg, arg);
              Console.WriteLine(msg, arg);
            }
            catch (Exception e)
            {
              m_f = null;
              Console.WriteLine("Exc in log:" + e);
            }
          }
        } 
    }

}