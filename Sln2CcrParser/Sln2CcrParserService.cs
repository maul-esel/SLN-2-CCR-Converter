using System;
using System.Text.RegularExpressions;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ChameleonCoder.Plugins;
using System.Xml;
using System.Windows.Forms;

namespace Sln2CcrParser
{
    [ChameleonCoder.CCPlugin]
    public class Sln2CcrParserService : IService
    {
        public string About
        {
            get { return "© 2011 maul.esel"; }
        }

        public string Author
        {
            get { return "maul.esel"; }
        }

        public string Description
        {
            get { return "converts *.sln files into *.ccr files"; }
        }

        public ImageSource Icon
        {
            get { return new BitmapImage(new Uri("pack://application:,,,/Sln2CcrParser;component/logo.png"))
                .GetAsFrozen() as ImageSource; }
        }

        public Guid Identifier
        {
            get { return new Guid("{f946cdaf-b912-4e16-8451-4e014ca58651}"); }
        }

        public bool IsBusy
        {
            get;
            private set;
        }

        public string Name
        {
            get { return "SLN 2 CCR parser"; }
        }

        public string Version
        {
            get { return "0.0.0.1"; }
        }

        public void Initialize()
        {
        }

        public void Shutdown()
        {
        }

        public void Execute()
        {
            using (var dialog = new OpenFileDialog() { Filter = "VS solution files (*.sln)|*.sln" })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var text = System.IO.File.ReadAllText(dialog.FileName);

                    if (Regex.IsMatch(text, @"^\s*Microsoft Visual Studio Solution File, Format Version 11.00\s*$", RegexOptions.ExplicitCapture | RegexOptions.Multiline))
                    {
                        var projects = ParseProjects(text);

                        var document = new XmlDocument();
                        document.LoadXml(initialXml);

                        var resources = (XmlElement)document.SelectSingleNode("/cc-resource-file/resources");
                        foreach (var project in projects)
                        {
                            var node = document.CreateElement("project");

                            node.SetAttribute("name", project.Item2);
                            node.SetAttribute("id", project.Item4.ToString("b"));

                            resources.AppendChild(node);

                            // todo: create resource-data elements & set "created at", "last modified", ...

                            // todo: parse project file
                        }
                        document.Save(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\test.ccr");
                    }
                    System.Windows.MessageBox.Show("complete!");
                }
            }
        }

        private static Tuple<Guid, string, string, Guid>[] ParseProjects(string text)
        {
            var matches = projectRegex.Matches(text);
            var projects = new Tuple<Guid, string, string, Guid>[matches.Count];

            int i = 0;
            foreach (Match match in matches)
            {
                projects[i] = Tuple.Create(Guid.Parse(match.Groups["type"].Value), match.Groups["name"].Value, match.Groups["file"].Value, Guid.Parse(match.Groups["id"].Value));
                i++;
            }

            return projects;
        }

        static string regex_guid = @"\{[a-zA-Z\d]{8}\-([a-zA-Z\d]{4}\-){3}[a-zA-Z\d]{12}}";
        static Regex projectRegex = new Regex(@"Project\(""(?<type>" + regex_guid + @")""\)\s*=\s*""(?<name>[^"",]+)"",\s*""(?<file>[^"",]+)"",\s*""(?<id>" + regex_guid + @")""[\s\r\n]*"
                        + @"(ProjectSection\(ProjectDependencies\)\s*=\s*postProject[\s\r\n]*"
                        + regex_guid + @"\s*=\s*" + regex_guid + @"[\s\r\n]*"
                        + @"EndProjectSection[\s\r\n]*)*"
                        + @"EndProject");

        static string initialXml = "<cc-resource-file><resources/><data/><settings/><references/></cc-resource-file>";
    }
}
