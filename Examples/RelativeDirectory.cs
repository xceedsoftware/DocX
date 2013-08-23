using System;
using System.IO;

namespace Examples
{
    class RelativeDirectory
    {
        // Author D. Bolton see http://cplus.about.com (c) 2010
        private DirectoryInfo _dirInfo;

        public string Dir
        {
            get
            {
                return _dirInfo.Name;
            }
        }

        public string Path
        {
            get { return _dirInfo.FullName; }
            set
            {
                try
                {
                    DirectoryInfo newDir = new DirectoryInfo(value);
                    _dirInfo = newDir;
                }
                catch
                {
                    // silent
                }
            }
        }
        public RelativeDirectory()
        {
            _dirInfo = new DirectoryInfo(Environment.CurrentDirectory);
        }

        public RelativeDirectory(string absoluteDir)
        {
            _dirInfo = new DirectoryInfo(absoluteDir);
        }

        public Boolean Up(int numLevels)
        {
            for (int i = 0; i < numLevels; i++)
            {
                DirectoryInfo tempDir = _dirInfo.Parent;
                if (tempDir != null)
                    _dirInfo = tempDir;
                else
                    return false;
            }
            return true;
        }

        public Boolean Up()
        {
            return Up(1);
        }

        public Boolean Down(string match)
        {
            DirectoryInfo[] dirs = _dirInfo.GetDirectories(match + '*');
            _dirInfo = dirs[0];
            return true;
        }

    }

}
