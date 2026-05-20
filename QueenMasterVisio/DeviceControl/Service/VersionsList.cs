using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Threading.Tasks;

namespace QueenMasterVisio.DeviceControl.Service
{
    public class VersionsList
    {
        public List<(string, float)> Versions { get; }
        public VersionsList(string versionsFilePath)
        {
            if (string.IsNullOrEmpty(versionsFilePath))
                return;
            try
            {
                var versionsList = new List<(string name, float version)>();
                string text = File.ReadAllText(versionsFilePath);
                foreach (string line in text.Split('\n'))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    string[] parts = line.Split('=');
                    if (parts.Length == 2 && float.TryParse(parts[1].Replace('.',',').Trim(), out float version))
                    {
                        versionsList.Add((parts[0].Trim(), version));
                    }
                }
                Versions = versionsList;
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
