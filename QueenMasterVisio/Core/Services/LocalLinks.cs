using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace QueenMasterVisio.Core.Services
{
    /// <summary>
    /// Класс выдает все локальные ссылки шаблонов, файлов, фигур
    /// </summary>
    public class LocalLinks
    {
        public string BasePath { get; }
        public string BaseDirectoryPath { get; }
        public string FileName { get; }
        public string FileNameWithoutEx { get; }
        public string MetafilesDir { get; }
        public string ChangeLogsDir { get; }
        public string QueenFigures { get; }

        public bool isLocalFile = true;

        public LocalLinks(string documentFullName)
        {
            if (string.IsNullOrEmpty(documentFullName))
                return;

            if (!(documentFullName.Contains("EscapeRoomDoctor") && documentFullName.Contains("Project")))
                return;

            int startIndex = documentFullName.IndexOf("EscapeRoomDoctor");
            if (startIndex == -1)
                return;


            string relativePath = documentFullName.Substring(startIndex).Replace('/', '\\');
            string userProfile = Environment.GetEnvironmentVariable("USERPROFILE");
            string basePath = System.IO.Path.Combine(userProfile, "OneDrive", relativePath);
            string directory = System.IO.Path.GetDirectoryName(basePath);

            if (!Directory.Exists(directory))
                return;

            BasePath = basePath;
            BaseDirectoryPath = directory;
            FileName = Path.GetFileName(basePath);
            FileNameWithoutEx = Path.GetFileNameWithoutExtension(basePath);

            //Папка для метафайлов
            string metafilesDir = System.IO.Path.Combine(directory, "Metafiles");
            if (!Directory.Exists(metafilesDir))
                Directory.CreateDirectory(metafilesDir);

            if (Directory.Exists(metafilesDir))
                MetafilesDir = metafilesDir;

            //Папка для чейнджлогов
            string changeLogsDir = System.IO.Path.Combine(metafilesDir, "ChangeLogs");
            if (!Directory.Exists(changeLogsDir))
                Directory.CreateDirectory(changeLogsDir);

            if (Directory.Exists(changeLogsDir))
                ChangeLogsDir = changeLogsDir;

            string QueenFiguresFilePath = System.IO.Path.Combine(userProfile, "OneDrive", "EscapeRoomDoctor\\Project\\!ШАБЛОН\\Фигуры\\Queen Figures.vssx");
            if (File.Exists(QueenFiguresFilePath))
            {
                QueenFigures = QueenFiguresFilePath;
            }

            isLocalFile = false;

        }
    }
}
