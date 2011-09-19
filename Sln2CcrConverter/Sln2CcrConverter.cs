using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml;
using ChameleonCoder.Plugins;

namespace Sln2CcrParser
{
    [ChameleonCoder.CCPlugin]
    public class Sln2CcrConverter : IService
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
            /*
             * TODO:
             * first show a window,
             * letting the user set some options,
             * opening (multiple) files,
             * evtl. also supporting reverse-conversion,
             * specifying the output file(s),
             * ...
             * 
             * It would be best if everything, even the resource types being created,
             * were customizable.
             */
            string file = null;

            using (var dialog = new OpenFileDialog() { Filter = "VS solution files (*.sln)|*.sln" })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    file = dialog.FileName;
                }
            }

            if (!string.IsNullOrWhiteSpace(file))
            {
                var dir = Path.GetDirectoryName(file);
                var text = File.ReadAllText(file);

                if (IsValidSLN(text))
                {
                    ccrDocument.LoadXml(initialXml);
                    var projects = ParseProjects(text);                    

                    /* todo: also allow only parsing of a *.csproj file, detect the difference */

                    foreach (var project in projects)
                    {
                        var node = ccrDocument.CreateElement("project"); // create the project definition
                        node.SetAttribute("name", project.Item2);
                        node.SetAttribute("id", project.Item4.ToString("b"));
                        ccrDocument.SelectSingleNode("/cc-resource-file/resources").AppendChild(node);

                        var data = ccrDocument.CreateElement("resource-data"); // create the resource-data element
                        data.SetAttribute("id", project.Item4.ToString("b"));
                        ccrDocument.SelectSingleNode("/cc-resource-file/data").AppendChild(data);

                        var created = ccrDocument.CreateElement("created"); // create the "created at" date
                        created.InnerText = DateTime.Now.ToString("yyyyMMddHHmmss");
                        data.AppendChild(created);

                        string full = Path.Combine(dir, project.Item3);
                        if (File.Exists(full))
                        {
                            var csproj = File.ReadAllText(full);
                            projDocument.LoadXml(csproj);

                            var xmlNsMan = new XmlNamespaceManager(projDocument.NameTable);
                            xmlNsMan.AddNamespace("vs", "http://schemas.microsoft.com/developer/msbuild/2003");

                            /*
                             * NOTE:
                             * to make this working ideally,
                             * some language-specific resource or RichContent types would be necessary.
                             * Plus some decisions. ;-)
                             * 
                             * For example, we could store assembly references as
                             *  - references to resources to be included in the file
                             *  - RichContent
                             *  - a resource property
                             *  - metadata
                             *  - ...
                             *  
                             *  We could store compiler information as
                             *  - metadata
                             *  - resource properties
                             *  - RichContent
                             *  - ...
                             *  
                             * =====================================
                             * 
                             * NOTE:
                             * when parsing resources, we should use the hierarchical structur in VS,
                             * so we need to create "Group" resources for folders.
                             * So we have to detect whether a resource is in the top directory or not,
                             * including support for multiple levels of groups.
                             * 
                             * Another thing is the hierarchy e.g. between a *.resx file and the Designer code,
                             * which should be covered, too.
                             */
                            foreach (XmlElement incl in projDocument.SelectNodes("/vs:Project/vs:ItemGroup/vs:Compile", xmlNsMan))
                            {
                                /* create code resources */
                            }

                            foreach (XmlElement res in projDocument.SelectNodes("/vs:Project/vs:ItemGroup/vs:EmbeddedResource", xmlNsMan))
                            {
                                /* create file resources */
                                var fileres = ccrDocument.CreateElement("file");
                                fileres.SetAttribute("path", res.GetAttribute("Include"));
                                fileres.SetAttribute("id", Guid.NewGuid().ToString("b"));
                                fileres.SetAttribute("name", Path.GetFileName(res.GetAttribute("Include")));

                                node.AppendChild(fileres);
                            }

                            foreach (XmlElement res in projDocument.SelectNodes("/vs:Project/vs:ItemGroup/vs:None", xmlNsMan))
                            {
                                /* create file resources */
                                var fileres = ccrDocument.CreateElement("file");
                                fileres.SetAttribute("path", res.GetAttribute("Include"));
                                fileres.SetAttribute("id", Guid.NewGuid().ToString("b"));
                                fileres.SetAttribute("name", Path.GetFileName(res.GetAttribute("Include")));

                                node.AppendChild(fileres);
                            }

                            foreach (XmlElement res in projDocument.SelectNodes("/vs:Project/vs:ItemGroup/vs:Reference", xmlNsMan))
                            {
                                /* create references to .NET assemblies (?) */
                            }

                            foreach (XmlElement res in projDocument.SelectNodes("/vs:Project/vs:ItemGroup/vs:Resource", xmlNsMan))
                            {
                                /* create file resources */
                                var fileres = ccrDocument.CreateElement("file");
                                fileres.SetAttribute("path", res.GetAttribute("Include"));
                                fileres.SetAttribute("id", Guid.NewGuid().ToString("b"));
                                fileres.SetAttribute("name", Path.GetFileName(res.GetAttribute("Include")));

                                node.AppendChild(fileres);
                            }

                            foreach (XmlElement res in projDocument.SelectNodes("/vs:Project/vs:ItemGroup/vs:Page", xmlNsMan))
                            {
                                /* create file resources or sth. else */
                                var fileres = ccrDocument.CreateElement("file");
                                fileres.SetAttribute("path", res.GetAttribute("Include"));
                                fileres.SetAttribute("id", Guid.NewGuid().ToString("b"));
                                fileres.SetAttribute("name", Path.GetFileName(res.GetAttribute("Include")));

                                node.AppendChild(fileres);
                            }
                        }

                        // todo: parse project file
                    }
                    ccrDocument.Save(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\test.ccr");
                }
                System.Windows.MessageBox.Show("complete!");
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

        #region validation

        private static bool IsValidSLN(string content)
        {
            return slnRegex.IsMatch(content);
        }
        static Regex slnRegex = new Regex(@"^\s*Microsoft Visual Studio Solution File, Format Version 11.00\s*$", RegexOptions.Multiline);

        #endregion

        static string initialXml = "<cc-resource-file><resources/><data/><settings/><references/></cc-resource-file>";

        private XmlDocument ccrDocument = new XmlDocument();

        private XmlDocument projDocument = new XmlDocument();
    }
}
