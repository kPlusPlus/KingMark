using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Web;
using System.ComponentModel;
using System.Net;
using System.Threading;
using MohammadDayyanCalendar;

namespace MohammadDayyan
{
    public partial class FormKingMark : Form
    {
        #region some of fields

        /// <summary>
        /// XML file name
        /// </summary>
        string temporaryFileName = "KingMarkXML.xml";

        /// <summary>
        /// stores current treeView selected node
        /// </summary>
        private TreeNode selected_node;

        /// <summary>
        /// For working with xml file
        /// </summary>
        XElement X_Element;

        private const string rootNodeName = "KingMark";
        private const string XmlDescription = "\r\nKingMark\r\nCreated by Mohammad Dayyan\r\nMy weblog : http://www.mds-soft.persianblog.ir/\r\n";
        private const string TempVariable = "TEMP";

        #endregion

        public FormKingMark()
        {
            InitializeComponent();

            //Adding Tray minimizer button beside of Minimize button
            //Reference : http://www.thecodeking.co.uk/2007/09/adding-caption-buttons-to-non-client.html
            /*
            IActiveMenu menu = ActiveMenu.GetInstance(this);
            ActiveButton button = new ActiveButton();/*
            button.Image = global::MohammadDayyan.Properties.Resources.trayButton;
            button.Name = "buttonTrayMinimizer";
            button.Click += new EventHandler(buttonTrayMinimizer_Click);
            button.BackColor = Color.LightBlue;            
            menu.Items.Add(button);
            */

        }
        

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region Form events

        private void FormKingMark_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.notifyIcon1.Visible = false;
        }

        private void FormKingMark_Resize(object sender, EventArgs e)
        {
            try
            {
                if (this.WindowState == FormWindowState.Minimized)
                {
                    createContextMenuStripNotifyIcon();//creates ContextMenuStrip of notify icon
                    this.Visible = false;
                    notifyIcon1.Visible = true;
                    notifyIcon1.ShowBalloonTip(1000);
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region messages

        private void error(ref StackFrame file_info, string errorMassage)
        {
            try
            {
                if (file_info.GetFileName() == null)
                    MessageBox.Show(this, "Exception : " + errorMassage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show(this, "File : " + file_info.GetFileName() + "\nLine : " + file_info.GetFileLineNumber().ToString() + "\nException : " + errorMassage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }

        private void successful(string title, string message)
        {
            try
            {
                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region FF bookmark file parser, converts FF bookmark file to List<string>

        /// <summary>
        /// Calculating numbers of spaces in the beginning of input string   
        /// </summary>
        /// <param name="str">A line that reads from FF bookmark file</param>
        /// <returns>number of spaces into the string</returns>
        private string spaceChar(string str)
        {
            string space = "";
            for (int i = 0; i < str.Length; i++)
                if (str[i] == ' ') space = " " + space;
                else break;
            return space;
        }

        /// <summary>
        /// Extracts 'H3' tags from FF bookmark file
        /// </summary>
        private string h3(string str)
        {
            string addDate = Regex.Match(str, "ADD_DATE=\"[^\"]*\"").ToString();
            string lastModified = Regex.Match(str, "LAST_MODIFIED=\"[^\"]*\"").ToString();

            str = Regex.Replace(str, @"<DT\>\<H3[^\>]*\>\s*", "");
            str = Regex.Replace(str, @"</H3\>\s*", "");

            str = spaceChar(str) + "NAME=\"" + str.Trim() + "\" " + addDate + " " + lastModified;
            return str;
        }

        /// <summary>
        /// Extracts 'a' tags from FF bookmark file
        /// </summary>
        private string a(string str)
        {
            string name = Regex.Replace(str, @"\<DT\>\s*\<A[^\>]*\>\s*", "", RegexOptions.IgnoreCase);
            name = Regex.Replace(name, @"\<\/A\>\s*", "", RegexOptions.IgnoreCase);
            name = Regex.Replace(name, "\"", "\\\"");
            name = spaceChar(name) + "NAME=\"" + name.Trim() + "\" ";

            string href = Regex.Match(str, "HREF=\"[^\"]*\"").ToString();

            string iconUrl = Regex.Match(str, "ICON_URI=\"[^\"]*\"").ToString();
            string icon = Regex.Match(str, "ICON=\"[^\"]*\"").ToString();

            string keyWords = Regex.Match(str, "SHORTCUTURL=\"[^\"]*\"").ToString();

            string lastModified = Regex.Match(str, "LAST_MODIFIED=\"[^\"]*\"").ToString();
            string addDate = Regex.Match(str, "ADD_DATE=\"[^\"]*\"").ToString();

            return name + " " + keyWords + " " + href + " " + addDate + " " + lastModified + " " + iconUrl + " " + icon + " ";
        }

        /// <summary>
        /// Checks input string to detect <DL> tags
        /// </summary>
        private bool isDL(string str)
        {
            return Regex.IsMatch(str, @"\<DL\>");
        }

        /// <summary>
        /// Checks input string to detect </DL> tags
        /// </summary>
        private bool isBDL(string str)
        {
            return Regex.IsMatch(str, @"\<\/DL\>");
        }

        /// <summary>
        /// Checking and extracting FF bookmark file's elements
        /// </summary>
        private string analyzer(string lineContent)
        {
            if (isDL(lineContent))
                return spaceChar(lineContent) + "<NODE>";

            else if (isBDL(lineContent))
                return spaceChar(lineContent) + "</NODE>";

            else if (Regex.IsMatch(lineContent, @"\<DT\>\<H3"))
                return h3(lineContent);

            else if (Regex.IsMatch(lineContent, @"\<DT\>\<A"))
                return a(lineContent);

            else if (Regex.IsMatch(lineContent, @"\<HR\>"))
                return "<HR>";

            else
                return "";
        }

        /// <summary>
        /// Parses list of FF Bookmark file's content
        /// </summary>
        /// <param name="InputList">
        /// content of FF Bookmark file
        /// each line is in each index of list
        /// </param>
        /// <returns>
        /// intermediate code to create XML from FF Bookmark file
        /// see 'trace.txt'
        /// </returns>
        private List<string> parse(List<string> InputList)
        {
            List<string> list = new List<string>();
            string temp = "", desc = "";

            try
            {
                for (int i = 0; i < InputList.Count; i++)
                {
                    //detects comments of elements
                    //برای تشخیص توضیحات 
                    if (Regex.IsMatch(InputList[i], @"\<DD\>"))
                    {
                        desc = Regex.Replace(InputList[i], @"\s*\<DD\>\s*", "").Trim();
                        while (i < InputList.Count)
                        {
                            i++;
                            //وقتی به <...> برخورد کنیم از حلقه خارج میشیم
                            if (Regex.IsMatch(InputList[i], @"\<[^\>]*>")) break;
                            desc += " \r\n" + InputList[i];
                        }
                        //بعد از تشخیص دادن توضیحات اون رشته رو با رشته ای جدید که شامل توضیحات است عوض می کنیم
                        //Last index => list.Count - 1
                        temp = list[list.Count - 1];
                        list.RemoveAt(list.Count - 1);
                        list.Add(temp + " DESC=\"" + desc + "\"");
                        desc = "";
                    }
                    //\\
                    temp = analyzer(InputList[i]);
                    if (temp != "")
                        list.Add(temp);
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }

            return list;
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region converting FF bookmark file to List and TreeView

        /// <summary>
        /// Creates "trace.txt" , intermediate code to crating XML file
        /// </summary>
        /// <param name="list">output of this.pars method</param>
        private void trace(List<string> list)
        {
            // Create an instance of StreamWriter to write text to a file.
            // The using statement also closes the StreamWriter.
            using (StreamWriter sw = new StreamWriter("trace.txt"))
            {
                foreach (var value in list)
                    sw.WriteLine(value);
                sw.Close();
            }
        }

        /// <summary>
        /// converts list of string to the XML file
        /// </summary>
        /// <param name="list">intermediate string code to crating XML file</param>
        private void ListToXml(List<string> list)
        {
            string name = "";
            string href = "";
            string icon = "";
            string iconUrl = "";
            string desc = "";
            string keyWords = "";
            string addDate = "";
            string lastModified = "";

            try
            {

                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Indent = true;
                //settings.NewLineOnAttributes = true;
                settings.Encoding = Encoding.UTF8;
                settings.ConformanceLevel = ConformanceLevel.Document;

                System.Xml.XmlWriter Writer;

                Writer = XmlWriter.Create(temporaryFileName, settings);
                Writer.WriteStartDocument();
                Writer.WriteComment(XmlDescription);
                Writer.WriteStartElement(rootNodeName);
                Writer.WriteAttributeString("ELEMENTS", list.Count.ToString());

                for (int i = 1; i < list.Count; i++)
                {
                    //Nodes separator
                    if (Regex.IsMatch(list[i], @"\<HR\>"))
                    {
                        //Writer.WriteStartElement("HR");
                        //Writer.WriteEndElement();
                        continue;
                    }
                    else if (Regex.IsMatch(list[i], @"\<NODE\>"))
                    {
                        //node's NAME
                        name = Regex.Match(list[i - 1].Trim(), "NAME=\"[^\"]*\"").ToString();
                        name = Regex.Replace(name, "NAME|\"|=", "");

                        //node's DESC
                        desc = Regex.Match(list[i - 1].Trim(), "DESC=\"[^\"]*\"").ToString();
                        desc = Regex.Replace(desc, "DESC|\"|=", "");

                        //node's ADD_DATE
                        addDate = Regex.Match(list[i - 1].Trim(), "ADD_DATE=\"[^\"]*\"").ToString();
                        addDate = Regex.Replace(addDate, "ADD_DATE|\"|=", "");

                        //node's LAST_MODIFIED
                        lastModified = Regex.Match(list[i - 1].Trim(), "LAST_MODIFIED=\"[^\"]*\"").ToString();
                        lastModified = Regex.Replace(lastModified, "LAST_MODIFIED|\"|=", "");
                        //\\

                        Writer.WriteStartElement("NODE" + i.ToString());
                        Writer.WriteAttributeString("NAME", name);
                        Writer.WriteAttributeString("ADD_DATE", addDate);
                        Writer.WriteAttributeString("LAST_MODIFIED", lastModified);
                        Writer.WriteAttributeString("DESC", desc);
                    }
                    else if (list[i].Trim() == "</NODE>")
                    {
                        if (i + 1 != list.Count)
                            Writer.WriteEndElement();
                    }
                    else if (!Regex.IsMatch(list[i + 1], @"\<NODE\>"))
                    {
                        //bookmark's name extracting
                        name = Regex.Match(list[i], "NAME=\"[^\"]*\"").ToString();
                        name = Regex.Replace(name, "NAME|\"|=", "");

                        //bookmark's address extracting
                        href = Regex.Match(list[i].Trim(), "HREF=\"[^\"]*\"").ToString();
                        href = Regex.Replace(href, "HREF|\"|=", "");

                        //bookmark's icon's URL extracting
                        iconUrl = Regex.Match(list[i].Trim(), "ICON_URI=\"[^\"]*\"").ToString();
                        iconUrl = Regex.Replace(iconUrl, "ICON_URI|\"|=", "");

                        //bookmark's icon extracting
                        icon = Regex.Match(list[i].Trim(), "ICON=\"[^\"]*\"").ToString();
                        icon = Regex.Replace(icon, "\"|ICON|=", "");

                        //bookmark's description extracting
                        desc = Regex.Match(list[i].Trim(), "DESC=\"[^\"]*\"").ToString();
                        desc = Regex.Replace(desc, "DESC|\"|=", "");

                        //bookmark's SHORTCUTURL extracting
                        keyWords = Regex.Match(list[i].Trim(), "SHORTCUTURL=\"[^\"]*\"").ToString();
                        keyWords = Regex.Replace(keyWords, "SHORTCUTURL|\"|=", "");

                        //ADD_DATE
                        addDate = Regex.Match(list[i].Trim(), "ADD_DATE=\"[^\"]*\"").ToString();
                        addDate = Regex.Replace(addDate, "ADD_DATE|\"|=", "");

                        //LAST_MODIFIED
                        lastModified = Regex.Match(list[i].Trim(), "LAST_MODIFIED=\"[^\"]*\"").ToString();
                        lastModified = Regex.Replace(lastModified, "LAST_MODIFIED|\"|=", "");

                        //\\
                        /////////////////////////////////////////////////////////////
                        //write above data in XML file
                        Writer.WriteStartElement("BOOKMARK" + i.ToString());
                        Writer.WriteAttributeString("NAME", name.Trim());
                        Writer.WriteAttributeString("HREF", href.Trim());
                        Writer.WriteAttributeString("SHORTCUTURL", keyWords.Trim());
                        Writer.WriteAttributeString("DESC", desc.Trim());
                        Writer.WriteAttributeString("ADD_DATE", addDate.Trim());
                        Writer.WriteAttributeString("LAST_MODIFIED", lastModified.Trim());
                        Writer.WriteAttributeString("ICON_URI", iconUrl.Trim());
                        Writer.WriteAttributeString("ICON", icon.Trim());

                        Writer.WriteEndElement();
                    }
                }
                //Writing close tag for the root element.
                Writer.WriteEndDocument();
                Writer.Flush();
                Writer.Close();

                X_Element = XElement.Load(temporaryFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);

            }
        }

        /// <summary>
        /// Converts XML file to treeView
        /// </summary>
        private void XmlToTree()
        {
            optimize_menuItem.Enabled = exportIE_favorites.Enabled = exportXML_menuItem.Enabled = exportHtml_menuItem.Enabled = true;
            this.treeViewBookMark.Enabled = true;
            this.XElementClopboard = null;
            this.TreeNodeClipboard = null;

            try
            {
                // SECTION 1. Create a DOM Document and load the XML data into it.
                XmlDocument dom = new XmlDocument();
                dom.Load(this.temporaryFileName);

                // SECTION 2. Initialize the TreeView control.
                treeViewBookMark.Nodes.Clear();
                treeViewBookMark.Nodes.Add(new TreeNode(dom.DocumentElement.Name));
                TreeNode tNode = new TreeNode(rootNodeName);
                tNode = treeViewBookMark.Nodes[0];
                tNode.Name = rootNodeName;

                // SECTION 3. Populate the TreeView with the DOM nodes.
                AddNode(dom.DocumentElement, tNode);

                treeViewBookMark.SelectedNode = treeViewBookMark.Nodes[0];
                selected_node = treeViewBookMark.Nodes[0];
                treeViewBookMark.SelectedNode.Expand();
            }
            catch (Exception ex)
            {
                textBoxIconURL.Enabled = textBoxDesc.Enabled = textBoxName.Enabled = buttonDelIcon.Enabled = textBoxKeyWords.Enabled = textBoxHref.Enabled = false;
                optimize_menuItem.Enabled = exportIE_favorites.Enabled = exportXML_menuItem.Enabled = exportHtml_menuItem.Enabled = false;
                this.treeViewBookMark.Enabled = false;

                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return;
            }
        }

        /// <summary>
        /// Creates treeView
        /// References : 
        /// ms-help://MS.MSDNQTR.v90.en/enu_kbnetframeworkkb/netframeworkkb/317597.htm
        /// http://support.microsoft.com/kb/317597
        /// </summary>
        private void AddNode(XmlNode inXmlNode, TreeNode inTreeNode)
        {
            try
            {
                XmlNode xNode;
                TreeNode tNode;
                XmlNodeList nodeList;
                int i;
                int j = -1;

                // Loop through the XML nodes until the leaf is reached.
                // Add the nodes to the TreeView during the looping process.
                if (inXmlNode.HasChildNodes)
                {
                    nodeList = inXmlNode.ChildNodes;
                    for (i = 0; i <= nodeList.Count - 1; i++)
                    {
                        if (inXmlNode.ChildNodes[i].Name.Equals("HR"))
                            continue;
                        j++;

                        xNode = inXmlNode.ChildNodes[i];
                        TreeNode newTreeNode;
                        newTreeNode = new TreeNode(xNode.Attributes["NAME"].Value);
                        newTreeNode.Name = xNode.Name;
                        //تعیین عکس هر گره
                        //assigns icon of each node -> this code needs more process
                        if (loadIcons_menuItem.Checked)
                        {
                            Image icon = createIcon(xNode.Name);
                            if (Regex.IsMatch(inXmlNode.ChildNodes[i].Name, "NODE"))
                            {
                                newTreeNode.ImageIndex = 0;
                                newTreeNode.SelectedImageIndex = 0;
                            }
                            else if (icon != null && Regex.IsMatch(inXmlNode.ChildNodes[i].Name, "BOOKMARK"))
                            {
                                imageList1.Images.Add(icon);
                                newTreeNode.ImageIndex = imageList1.Images.Count - 1;
                                newTreeNode.SelectedImageIndex = imageList1.Images.Count - 1;
                            }
                            else
                            {
                                newTreeNode.ImageIndex = 1;
                                newTreeNode.SelectedImageIndex = 1;
                            }
                        }
                        else
                        {
                            if (Regex.IsMatch(inXmlNode.ChildNodes[i].Name, "NODE"))
                            {
                                newTreeNode.ImageIndex = 0;
                                newTreeNode.SelectedImageIndex = 0;
                            }
                            else if (Regex.IsMatch(inXmlNode.ChildNodes[i].Name, "BOOKMARK"))
                            {
                                newTreeNode.ImageIndex = 1;
                                newTreeNode.SelectedImageIndex = 1;
                            }
                        }
                        //\\
                        inTreeNode.Nodes.Add(newTreeNode);
                        tNode = inTreeNode.Nodes[j];
                        AddNode(xNode, tNode);
                    }
                }
                else
                {
                    // Here you need to pull the data from the XmlNode based on the
                    // type of node, whether attribute values are required, and so forth.                    
                    try
                    {
                        inTreeNode.Text = inXmlNode.Attributes["NAME"].Value.Trim();
                    }
                    catch
                    {
                        inTreeNode.Text = rootNodeName;
                    }
                }
            }
            catch (Exception ex)
            {
                textBoxIconURL.Enabled = textBoxDesc.Enabled = textBoxName.Enabled = buttonDelIcon.Enabled = textBoxKeyWords.Enabled = textBoxHref.Enabled = false;
                optimize_menuItem.Enabled = exportIE_favorites.Enabled = exportXML_menuItem.Enabled = exportHtml_menuItem.Enabled = false;
                treeViewBookMark.Nodes.Clear();
                this.treeViewBookMark.Enabled = false;

                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return;
            }
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////        

        #region XML operations

        /// <summary>
        /// Extracts address of a bookmark from KingMark XML file
        /// </summary>
        /// <param name="nodeName">name of node to extracting address</param>
        /// <returns>Address of bookmark</returns>
        private string GetHref(string nodeName)
        {
            try
            {
                //XDocument Xdom = XDocument.Load(FileName);
                var href = from node in X_Element.Descendants(nodeName)
                           //where node.Attribute("ID").Value == id
                           select node.Attribute("HREF").Value;

                return href.First();
            }
            catch (Exception)
            {
                return "";
            }
        }

        /// <summary>
        /// creates icon of node
        /// </summary>
        private Image createIcon(string nodeName)
        {
            try
            {
                var query = from node in X_Element.Descendants(nodeName)
                            select node.Attribute("ICON").Value;

                string base64_string = query.First();
                base64_string = Regex.Replace(base64_string, @"^[^,]*,", "");

                return Base64StringToImage(base64_string);
            }
            catch
            {
                return null;
            }
        }

        //http://www.codeproject.com/KB/GDI-plus/image-base-64-converter.aspx

        private string ImageToBase64String(Image imageData, ImageFormat format)
        {
            MemoryStream memory = new MemoryStream();
            imageData.Save(memory, format);
            string base64 = System.Convert.ToBase64String(memory.ToArray());
            memory.Close();
            memory.Dispose();
            return base64;
        }

        private Image Base64StringToImage(string base64ImageString)
        {
            if (base64ImageString == "") return null;
            byte[] b;
            //تصحیح کردن رشته ورودی
            //Corrects input string length 
            bool error = true;
            if (base64ImageString[base64ImageString.Length - 1] != '=')
                while (error)
                {
                    base64ImageString += "I";
                    try
                    {
                        b = Convert.FromBase64String(base64ImageString);
                        error = false;
                    }
                    catch
                    {
                        error = true;
                    }
                }
            b = Convert.FromBase64String(base64ImageString);
            MemoryStream ms = new System.IO.MemoryStream(b);
            Image img = System.Drawing.Image.FromStream(ms);
            return img;
        }

        //\\

        /// <summary>
        /// Extracts description of a bookmark from KingMark XML file
        /// </summary>
        private string GetDescription(string nodeName)
        {
            try
            {
                //XDocument Xdom = XDocument.Load(FileName);
                var desc = from node in X_Element.Descendants(nodeName)
                           select node.Attribute("DESC").Value;

                return (desc.First());
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Extracts keyword of a bookmark from XML file
        /// </summary>
        private string GetKeyword(string nodeName)
        {
            try
            {
                var keyword = from node in X_Element.Descendants(nodeName)
                              select node.Attribute("SHORTCUTURL").Value;

                return (keyword.First());
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Gets ADD_DATE of a node
        /// </summary>
        private Double GetAddDate(string nodeName)
        {
            try
            {
                var add_date = from node in X_Element.Descendants(nodeName)
                               select node.Attribute("ADD_DATE").Value;

                return (Double.Parse(add_date.First()));
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Gets LAST_MODIFIED of a node
        /// </summary>
        private Double GetLastModified(string nodeName)
        {
            try
            {
                var last_Modified = from node in X_Element.Descendants(nodeName)
                                    select node.Attribute("LAST_MODIFIED").Value;

                return (Double.Parse(last_Modified.First()));
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Gets ICON_URI of a node
        /// </summary>
        private string GetIconUri(string nodeName)
        {
            try
            {
                var iconUrl = from node in X_Element.Descendants(nodeName)
                              select node.Attribute("ICON_URI").Value;

                return (iconUrl.First());
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Deletes some nodes from XML file
        /// </summary>
        private void deleteNodes(string nodeName)
        {
            try
            {
                var SelectedNodes = (from node in X_Element.Descendants(nodeName)
                                     select node).First();

                int max = Int32.Parse(X_Element.Attribute("ELEMENTS").Value);
                max -= (from node in SelectedNodes.Descendants() select node).Count();

                SelectedNodes.Remove();
                treeViewBookMark.Nodes.Remove(selected_node);

                X_Element.Attribute("ELEMENTS").Value = max.ToString();
                X_Element.Save(this.temporaryFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        /// <summary>
        /// Clears icon's attribute value
        /// </summary>
        private void clearIcon(string nodeName)
        {
            try
            {
                var SelectedNode = (from node in X_Element.Descendants(nodeName)
                                    select node).First();
                SelectedNode.SetAttributeValue("ICON", "");
                X_Element.Save(temporaryFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        /// <summary>
        /// Reloads the icon from its website and changes the value of it in XML file
        /// </summary>
        private void reloadIcon(string nodeName)
        {
            try
            {
                if (Regex.IsMatch(nodeName, "NODE|" + rootNodeName)) return;

                var selectedNode = (from node in this.X_Element.Descendants(nodeName)
                                    select node).First();
                toolStripStatusLabel1.Text = "Reloading " + selectedNode.Attribute("NAME").Value + " icon ...";

                BackgroundWorker backgroundWorker = new BackgroundWorker();
                backgroundWorker.DoWork += new DoWorkEventHandler(backgroundWorker_DoWork);
                backgroundWorker.RunWorkerAsync(nodeName);
            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = ex.Message;
            }
        }

        private void backgroundWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                var selectedNode = (from node in this.X_Element.Descendants(e.Argument.ToString())
                                    select node).First();
                string ICON_URI = selectedNode.Attribute("ICON_URI").Value;
                string Base64String = readIcon(ICON_URI);
                if (Base64String == "")
                {
                    toolStripStatusLabel1.Text = "KingMark couldn't find icon uri";
                    return;
                }
                Image newIcon = Base64StringToImage(Base64String);
                //changing Icon size
                if (newIcon.Height > 16 && newIcon.Width > 16)
                {
                    newIcon = newIcon.GetThumbnailImage(16, 16, null, IntPtr.Zero);
                    Base64String = ImageToBase64String(newIcon, ImageFormat.Png/*GetImageFormat()*/);
                }
                //\\
                selectedNode.SetAttributeValue("ICON", "data:image/x-icon;base64," + Base64String);
                if (newIcon != null)
                    this.pictureBoxIcon.Image = newIcon; //for prevention from occurring Exception on this line this.pictureBoxIcon.SizeMode should be PictureBoxSizeMode.Normal;
                if (this.loadIcons_menuItem.Checked)
                {
                    int index = this.treeViewBookMark.Nodes.Find(e.Argument.ToString(), true).First().ImageIndex;
                    if (index != 0 && index != 1)//because 0 & 1 are KingMark's icons and don't have to change.
                    {
                        this.imageList1.Images[index] = newIcon;
                        this.treeViewBookMark.Nodes.Find(e.Argument.ToString(), true).First();
                    }
                }
                this.X_Element.Save(temporaryFileName);
                toolStripStatusLabel1.Text = "Icon of " + selectedNode.Attribute("NAME").Value + " realoded successfully";
            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = ex.Message;
            }
        }

        /// <summary>
        /// A function for changing image size
        /// I didn't use it
        /// </summary>
        private Image resizeImage(Image img, int width, int height)
        {
            Bitmap b = new Bitmap(width, height);
            Graphics g = Graphics.FromImage((Image)b);

            g.DrawImage(img, 0, 0, width, height);
            g.Dispose();

            return (Image)b;
        }

        private ImageFormat GetImageFormat(Image image)
        {
            if (image.RawFormat.Equals(ImageFormat.Bmp)) return ImageFormat.Bmp;
            else if (image.RawFormat.Equals(ImageFormat.Emf)) return ImageFormat.Emf;
            else if (image.RawFormat.Equals(ImageFormat.Exif)) return ImageFormat.Exif;
            else if (image.RawFormat.Equals(ImageFormat.Gif)) return ImageFormat.Gif;
            else if (image.RawFormat.Equals(ImageFormat.Icon)) return ImageFormat.Icon;
            else if (image.RawFormat.Equals(ImageFormat.Jpeg)) return ImageFormat.Jpeg;
            else if (image.RawFormat.Equals(ImageFormat.MemoryBmp)) return ImageFormat.MemoryBmp;
            else if (image.RawFormat.Equals(ImageFormat.Png)) return ImageFormat.Png;
            else if (image.RawFormat.Equals(ImageFormat.Tiff)) return ImageFormat.Tiff;
            else if (image.RawFormat.Equals(ImageFormat.Wmf)) return ImageFormat.Wmf;
            else return ImageFormat.Gif;
        }

        private string readIcon(string uri)
        {
            try
            {
                WebClient webClient = new WebClient();
                byte[] b = webClient.DownloadData(uri);
                string str = Convert.ToBase64String(b);
                return str;
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Updates XML file and treeView with new data
        /// </summary>
        /// <param name="nodeName">Name of changed node</param>
        private void update(string nodeName)
        {
            try
            {
                if (textBoxHref.Text.Trim() == "" && !Regex.IsMatch(nodeName, @"NODE|" + rootNodeName))
                    textBoxHref.Text = "Location";
                else if (textBoxName.Text.Trim() == "")
                    textBoxName.Text = "Name";

                MDCalendar calendar = new MDCalendar();
                DateTime date = DateTime.Now;
                TimeZone time = TimeZone.CurrentTimeZone;
                TimeSpan difference = time.GetUtcOffset(date);
                uint currentTime = calendar.Time() + (uint)difference.TotalSeconds;

                var Selected_Node = (from node in X_Element.Descendants(nodeName)
                                     select node).First();

                if (textBoxName.Text.Trim() != "")
                    Selected_Node.SetAttributeValue("NAME", textBoxName.Text.Trim());
                if (textBoxHref.Text.Trim() != "")
                    Selected_Node.SetAttributeValue("HREF", textBoxHref.Text.Trim());

                Selected_Node.SetAttributeValue("DESC", textBoxDesc.Text.Trim());
                Selected_Node.SetAttributeValue("LAST_MODIFIED", currentTime.ToString());
                if (!Regex.IsMatch(nodeName, @"NODE|" + rootNodeName))
                {
                    Selected_Node.SetAttributeValue("ICON_URI", textBoxIconURL.Text.Trim());
                    Selected_Node.SetAttributeValue("SHORTCUTURL", textBoxKeyWords.Text.Trim());
                }

                X_Element.Save(temporaryFileName);

                treeViewBookMark.SelectedNode = selected_node;
                treeViewBookMark.SelectedNode.Text = textBoxName.Text.Trim();
            }
            catch { }
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region contextMenuStripTreeView

        /// <summary>
        /// stores a XElement item
        /// </summary>
        private XElement XElementClopboard = null;

        /// <summary>
        /// stores a TreeNode item
        /// </summary>
        private TreeNode TreeNodeClipboard = null;

        /// <summary>
        /// changes the name of nodes in XElementClopboard or TreeNodeClipboard
        /// I did it to preventing conflict  
        /// </summary>
        private XElement changeName(XElement xElement, ref int maxElements)
        {
            maxElements++;
            XElement newXElement = null;

            if (Regex.IsMatch(xElement.Name.ToString(), "NODE"))
                newXElement = new XElement("NODE" + maxElements.ToString(),
                                  new XAttribute("NAME", xElement.Attribute("NAME").Value),
                                  new XAttribute("ADD_DATE", xElement.Attribute("ADD_DATE").Value),
                                  new XAttribute("LAST_MODIFIED", xElement.Attribute("LAST_MODIFIED").Value),
                                  new XAttribute("DESC", xElement.Attribute("DESC").Value));

            else
                newXElement = new XElement("BOOKMARK" + maxElements.ToString(),
                                new XAttribute("NAME", xElement.Attribute("NAME").Value),
                                new XAttribute("HREF", xElement.Attribute("HREF").Value),
                                new XAttribute("SHORTCUTURL", xElement.Attribute("SHORTCUTURL").Value),
                                new XAttribute("DESC", xElement.Attribute("DESC").Value),
                                new XAttribute("ADD_DATE", xElement.Attribute("ADD_DATE").Value),
                                new XAttribute("LAST_MODIFIED", xElement.Attribute("LAST_MODIFIED").Value),
                                new XAttribute("ICON_URI", xElement.Attribute("ICON_URI").Value),
                                new XAttribute("ICON", xElement.Attribute("ICON").Value));

            foreach (XElement element in xElement.Elements())
            {
                if (element.Elements().Count() > 0)
                    newXElement.Add(changeName(element, ref maxElements));
                else
                {
                    maxElements++;
                    if (Regex.IsMatch(element.Name.ToString(), "NODE"))
                        newXElement.Add(new XElement("NODE" + maxElements.ToString(),
                                            new XAttribute("NAME", element.Attribute("NAME").Value),
                                            new XAttribute("ADD_DATE", element.Attribute("ADD_DATE").Value),
                                            new XAttribute("LAST_MODIFIED", element.Attribute("LAST_MODIFIED").Value),
                                            new XAttribute("DESC", element.Attribute("DESC").Value)));
                    else
                        newXElement.Add(new XElement("BOOKMARK" + maxElements.ToString(),
                                            new XAttribute("NAME", element.Attribute("NAME").Value),
                                            new XAttribute("HREF", element.Attribute("HREF").Value),
                                            new XAttribute("SHORTCUTURL", element.Attribute("SHORTCUTURL").Value),
                                            new XAttribute("DESC", element.Attribute("DESC").Value),
                                            new XAttribute("ADD_DATE", element.Attribute("ADD_DATE").Value),
                                            new XAttribute("LAST_MODIFIED", element.Attribute("LAST_MODIFIED").Value),
                                            new XAttribute("ICON_URI", element.Attribute("ICON_URI").Value),
                                            new XAttribute("ICON", element.Attribute("ICON").Value)));
                }
            }
            return newXElement;
        }
        private TreeNode changeName(TreeNode treeNode, ref int maxElements)
        {
            TreeNode newTreeNode = null;

            if (Regex.IsMatch(treeNode.Name, "NODE"))
            {
                newTreeNode = new TreeNode(treeNode.Text, treeNode.ImageIndex, treeNode.SelectedImageIndex);
                newTreeNode.Name = "NODE" + ++maxElements;
            }
            else if (Regex.IsMatch(treeNode.Name, "BOOKMARK"))
            {
                newTreeNode = new TreeNode(treeNode.Text, treeNode.ImageIndex, treeNode.SelectedImageIndex);
                newTreeNode.Name = "BOOKMARK" + ++maxElements;
            }

            for (int i = 0; i < treeNode.Nodes.Count; i++)
            {
                if (treeNode.Nodes[i].Nodes.Count > 0)
                    newTreeNode.Nodes.Add(changeName(treeNode.Nodes[i], ref maxElements));
                else
                {
                    if (Regex.IsMatch(treeNode.Nodes[i].Name, "NODE"))
                    {
                        ++maxElements;
                        newTreeNode.Nodes.Add("NODE" + maxElements, treeNode.Nodes[i].Text, treeNode.Nodes[i].ImageIndex, treeNode.Nodes[i].SelectedImageIndex);
                    }
                    else if (Regex.IsMatch(treeNode.Nodes[i].Name, "BOOKMARK"))
                    {
                        ++maxElements;
                        newTreeNode.Nodes.Add("BOOKMARK" + maxElements, treeNode.Nodes[i].Text, treeNode.Nodes[i].ImageIndex, treeNode.Nodes[i].SelectedImageIndex);
                    }
                }
            }
            return newTreeNode;
        }

        /// <summary>
        /// Copys treeView's item to the Clipboard
        /// </summary>
        private void copyToClipboard(string nodeName)
        {
            try
            {
                XElement elements = (from node in X_Element.Descendants(nodeName)
                                     select node).First();
                this.XElementClopboard = elements;
                this.TreeNodeClipboard = treeViewBookMark.SelectedNode;
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        /// <summary>
        /// Cuts treeView's item to the Clipboard
        /// </summary>
        private void cutToClipboard(string nodeName)
        {
            try
            {
                XElement elements = (from node in X_Element.Descendants(nodeName)
                                     select node).First();
                this.XElementClopboard = elements;
                this.TreeNodeClipboard = treeViewBookMark.SelectedNode;
                deleteNodes(elements.Name.ToString());
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        /// <summary>
        /// Pastes Clipboard's data into the treeView
        /// </summary>
        private void pasteClipboard()
        {
            try
            {
                XElement newXmlNode = this.XElementClopboard;
                TreeNode newTreeNode = (TreeNode)this.TreeNodeClipboard.Clone();

                //عوض کردن نام گره جدید 
                //تا با بقیه گره ها تداخلی نداشته باشد
                int max_elements = Int32.Parse(this.X_Element.Attribute("ELEMENTS").Value);
                newXmlNode = changeName(newXmlNode, ref max_elements);

                max_elements = Int32.Parse(this.X_Element.Attribute("ELEMENTS").Value);
                newTreeNode = changeName(newTreeNode, ref max_elements);

                this.treeViewBookMark.CollapseAll();

                XElement destinationElement;
                //اگر گره مقصد یه فولدر بود ما بوک مارک ها را داخل آن قرار می دهیم 
                //در غیر این صورت بوک مارک ها را داخل فولدر پدر قرار می دهیم
                if (Regex.IsMatch(this.selected_node.Name, rootNodeName))
                {
                    this.X_Element.Add(newXmlNode);
                    this.selected_node.Nodes.Add(newTreeNode);
                    this.treeViewBookMark.SelectedNode = this.selected_node = this.treeViewBookMark.Nodes.Find(newXmlNode.Name.ToString(), true).First();
                    this.selected_node.Collapse();
                }
                else if (Regex.IsMatch(this.selected_node.Name, "NODE"))
                {
                    destinationElement = (from x in X_Element.Descendants(selected_node.Name)
                                          select x).First();
                    destinationElement.Add(newXmlNode);
                    this.selected_node.Nodes.Add(newTreeNode);
                    this.treeViewBookMark.SelectedNode = this.selected_node = this.treeViewBookMark.Nodes.Find(newXmlNode.Name.ToString(), true).First();
                    this.selected_node.Collapse();
                }
                else if (Regex.IsMatch(this.selected_node.Name, "BOOKMARK"))
                {
                    destinationElement = (from x in X_Element.Descendants(selected_node.Name)
                                          select x).First().Parent;
                    var child = destinationElement.Element(selected_node.Name);
                    child.AddAfterSelf(newXmlNode);
                    this.selected_node.Parent.Nodes.Add(newTreeNode);
                    this.treeViewBookMark.SelectedNode = this.selected_node = this.treeViewBookMark.Nodes.Find(newXmlNode.Name.ToString(), true).First();
                    this.selected_node.Collapse();
                }
                this.X_Element.Attribute("ELEMENTS").Value = max_elements.ToString();
                this.X_Element.Save(this.temporaryFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);

            }
        }

        private void copyItem_Click(object sender, EventArgs e)
        {
            copyToClipboard(selected_node.Name);
        }

        private void cutItem_Click(object sender, EventArgs e)
        {
            cutToClipboard(selected_node.Name);
        }

        private void deleteItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure ?", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                deleteNodes(this.selected_node.Name);
        }

        private void pasteItem_Click(object sender, EventArgs e)
        {
            pasteClipboard();
        }

        private DirectoryInfo chooseNewFolderName(string fullPath)
        {
            try
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(fullPath);
                if (directoryInfo.Exists)
                    for (int i = 0; i < 10000; i++)
                        if (Directory.Exists(fullPath + i.ToString()))
                            continue;
                        else
                        {
                            directoryInfo = new DirectoryInfo(fullPath + i.ToString());
                            break;
                        }
                return directoryInfo;
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return null;
            }
        }

        private string chooseNewFileName(string fullName)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(fullName);
                if (fileInfo.Exists)
                    for (int i = 0; i < 10000; i++)
                        if (File.Exists(fileInfo.Directory + "\\" + Regex.Replace(fileInfo.Name, @"\.url", "", RegexOptions.IgnoreCase) + i.ToString() + ".url"))
                            continue;
                        else
                        {
                            fileInfo = new FileInfo(fileInfo.Directory + "\\" + Regex.Replace(fileInfo.Name, @"\.url", "", RegexOptions.IgnoreCase) + i.ToString() + ".url");
                            break;
                        }
                return fileInfo.FullName;
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return null;
            }
        }

        private void exportForIEItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                {
                    string rootFolder = rootNodeName;
                    try
                    {
                        rootFolder = (from n in X_Element.Descendants(this.treeViewBookMark.SelectedNode.Name)
                                      select n).First().Attribute("NAME").Value;
                    }
                    catch { }

                    if (this.treeViewBookMark.SelectedNode.Name == rootNodeName)
                    {
                        rootFolder = Regex.Replace(HttpUtility.HtmlDecode(rootFolder), @"http:\/\/|ftp:\/\/", "");
                        rootFolder = Regex.Replace(rootFolder, @"\/|//|:|\*|\?|\<|\>|\|" + "|\"", "");
                        if (rootFolder.Length > 260) rootFolder = rootFolder.Substring(0, 260);

                        directoryInfo = new DirectoryInfo(folderBrowserDialog1.SelectedPath + "//" + rootFolder);
                        Directory.CreateDirectory(directoryInfo.FullName);

                        foreach (var element in X_Element.Elements())
                            exportIE_favorite(element.Name.ToString());
                    }
                    else
                    {
                        var elements = from el in X_Element.Descendants(this.treeViewBookMark.SelectedNode.Name)
                                       select el;

                        if (elements.Elements().Count() == 0)
                        {
                            directoryInfo = new DirectoryInfo(folderBrowserDialog1.SelectedPath);
                            exportIE_favorite(elements.First().Name.ToString());
                        }
                        else
                        {
                            rootFolder = Regex.Replace(HttpUtility.HtmlDecode(rootFolder), @"http:\/\/|ftp:\/\/", "");
                            rootFolder = Regex.Replace(rootFolder, @"\/|//|:|\*|\?|\<|\>|\|" + "|\"", "");
                            if (rootFolder.Length > 250) rootFolder = rootFolder.Substring(0, 250);

                            directoryInfo = new DirectoryInfo(folderBrowserDialog1.SelectedPath + "//" + rootFolder);
                            //choosing new name for the folder
                            if (!overWriteFilesItem.Checked && directoryInfo.Exists)
                                directoryInfo = chooseNewFolderName(this.directoryInfo.FullName);

                            Directory.CreateDirectory(directoryInfo.FullName);

                            foreach (var element in elements.Elements())
                                exportIE_favorite(element.Name.ToString());
                        }
                    }
                    successful("Finish", "All data saved successfully");
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return;
            }
        }

        /// <summary>
        /// Navigates XML file and creates *.url(IE favorite's extension) files
        /// </summary>
        /// <param name="xElement_Name">Name of current element</param>
        private void exportIE_favorite(string xElement_Name)
        {
            try
            {
                var xElement = from n in X_Element.Descendants(xElement_Name) select n;

                string fileOrFolder_Name = Regex.Replace(HttpUtility.HtmlDecode(xElement.First().Attribute("NAME").Value), @"http:\/\/|ftp:\/\/", "");
                fileOrFolder_Name = Regex.Replace(fileOrFolder_Name, @"\/|//|:|\*|\?|\<|\>|\|" + "|\"", "");

                if (fileOrFolder_Name.Length > 250) fileOrFolder_Name = fileOrFolder_Name.Substring(0, 250);

                if (xElement.First().Attribute("NAME").Value == "Recently Bookmarked"
                    || xElement.First().Attribute("NAME").Value == "Recent Tags")
                    return;

                if (Regex.IsMatch(xElement.First().Name.ToString(), "NODE|" + rootNodeName))
                {
                    directoryInfo = new DirectoryInfo(directoryInfo.FullName + "\\" + fileOrFolder_Name);
                    //choosing the new name for the folder
                    if (!overWriteFilesItem.Checked && directoryInfo.Exists)
                        directoryInfo = chooseNewFolderName(directoryInfo.FullName);
                    Directory.CreateDirectory(directoryInfo.FullName);
                    foreach (var element in xElement.Elements())
                        exportIE_favorite(element.Name.ToString());
                    directoryInfo = directoryInfo.Parent;
                }
                else if (Regex.IsMatch(xElement.First().Name.ToString(), "BOOKMARK"))
                {
                    fileOrFolder_Name = this.directoryInfo.FullName + "\\" + fileOrFolder_Name + ".url";

                    if (!overWriteFilesItem.Checked && File.Exists(fileOrFolder_Name))
                        fileOrFolder_Name = chooseNewFileName(fileOrFolder_Name);
                    else if (File.Exists(fileOrFolder_Name))
                        File.Delete(fileOrFolder_Name);

                    using (StreamWriter sw = new StreamWriter(fileOrFolder_Name, true))
                    {
                        sw.Write(@"[InternetShortcut]" + Environment.NewLine
                                + "URL=" + xElement.First().Attribute("HREF").Value);
                        if (xElement.First().Attribute("ICON_URI").Value != "" && xElement.First().Attribute("ICON_URI").Value != null)
                            sw.Write(Environment.NewLine +
                                "IconFile=" + xElement.First().Attribute("ICON_URI").Value + Environment.NewLine
                                + "IconIndex=1");
                        sw.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return;
            }
        }

        private void exportForFFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                try
                {
                    XElement xElement = (from el in this.X_Element.Descendants(this.selected_node.Name)
                                         select el).First();
                    createFFBookmark(xElement);
                }
                catch
                {
                    createFFBookmark(this.X_Element);
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void newFolderItem_Click(object sender, EventArgs e)
        {
            createNewFolder(this.selected_node.Name);
        }

        private void createNewFolder(string nodeName)
        {
            try
            {
                int max_elements = Int32.Parse(this.X_Element.Attribute("ELEMENTS").Value);
                max_elements++;

                MDCalendar calendar = new MDCalendar();
                DateTime date = DateTime.Now;
                TimeZone time = TimeZone.CurrentTimeZone;
                TimeSpan difference = time.GetUtcOffset(date);
                uint currentTime = calendar.Time() + (uint)difference.TotalSeconds;

                XElement newFolder_element = new XElement("NODE" + max_elements,
                                                new XAttribute("NAME", "New Folder"),
                                                new XAttribute("ADD_DATE", currentTime.ToString()),
                                                new XAttribute("LAST_MODIFIED", currentTime.ToString()),
                                                new XAttribute("DESC", ""));

                if (Regex.IsMatch(this.selected_node.Name, rootNodeName))
                {
                    this.X_Element.Add(newFolder_element);

                    this.selected_node.Nodes.Add("NODE" + max_elements.ToString(), "New Folder", 0, 0);
                    this.selected_node.Expand();
                }
                else if (Regex.IsMatch(this.selected_node.Name, "NODE"))
                {
                    var destinationElement = (from xNode in X_Element.Descendants(nodeName)
                                              select xNode).First();
                    destinationElement.Add(newFolder_element);

                    this.selected_node.Nodes.Add("NODE" + max_elements.ToString(), "New Folder", 0, 0);
                    this.selected_node.Expand();
                }
                else if (Regex.IsMatch(this.selected_node.Name, "BOOKMARK"))
                {
                    var destinationElement = (from xNode in X_Element.Descendants(selected_node.Name)
                                              select xNode).First().Parent;
                    var child = destinationElement.Element(selected_node.Name);
                    child.AddAfterSelf(newFolder_element);

                    this.selected_node.Parent.Nodes.Add("NODE" + max_elements.ToString(), "New Folder", 0, 0);
                    this.selected_node.Parent.Expand();
                }

                this.X_Element.Attribute("ELEMENTS").Value = max_elements.ToString();
                this.X_Element.Save(this.temporaryFileName);

                this.treeViewBookMark.SelectedNode = this.selected_node = this.treeViewBookMark.Nodes.Find("NODE" + max_elements.ToString(), true).First();

            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return;
            }
        }

        private void createNewBookmark(string nodeName)
        {
            try
            {
                int max_elements = Int32.Parse(this.X_Element.Attribute("ELEMENTS").Value);
                max_elements++;

                MDCalendar calendar = new MDCalendar();
                DateTime date = DateTime.Now;
                TimeZone time = TimeZone.CurrentTimeZone;
                TimeSpan difference = time.GetUtcOffset(date);
                uint currentTime = calendar.Time() + (uint)difference.TotalSeconds;

                XElement newFolder_element = new XElement("BOOKMARK" + max_elements.ToString(),
                                                new XAttribute("NAME", "Name of new Bookmark"),
                                                new XAttribute("HREF", "New location"),
                                                new XAttribute("SHORTCUTURL", ""),
                                                new XAttribute("DESC", ""),
                                                new XAttribute("ADD_DATE", currentTime.ToString()),
                                                new XAttribute("LAST_MODIFIED", currentTime.ToString()),
                                                new XAttribute("ICON_URI", ""),
                                                new XAttribute("ICON", ""));

                if (Regex.IsMatch(this.selected_node.Name, rootNodeName))
                {
                    this.X_Element.Add(newFolder_element);

                    this.selected_node.Nodes.Add("BOOKMARK" + max_elements.ToString(), "New Bookmark", 1, 1);
                    this.selected_node.Expand();
                }
                else if (Regex.IsMatch(this.selected_node.Name, "NODE"))
                {
                    var destinationElement = (from xNode in X_Element.Descendants(nodeName)
                                              select xNode).First();
                    destinationElement.Add(newFolder_element);

                    this.selected_node.Nodes.Add("BOOKMARK" + max_elements.ToString(), "New Bookmark", 1, 1);
                    this.selected_node.Expand();
                }
                else if (Regex.IsMatch(this.selected_node.Name, "BOOKMARK"))
                {
                    var destinationElement = (from xNode in X_Element.Descendants(selected_node.Name)
                                              select xNode).First().Parent;
                    var child = destinationElement.Element(selected_node.Name);
                    child.AddAfterSelf(newFolder_element);

                    this.selected_node.Parent.Nodes.Add("BOOKMARK" + max_elements.ToString(), "New Bookmark", 1, 1);
                    this.selected_node.Parent.Expand();
                }

                this.X_Element.Attribute("ELEMENTS").Value = max_elements.ToString();
                this.X_Element.Save(this.temporaryFileName);

                this.treeViewBookMark.SelectedNode = this.selected_node = this.treeViewBookMark.Nodes.Find("BOOKMARK" + max_elements.ToString(), true).First();

            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return;
            }
        }

        private void newBookmarkItem_Click(object sender, EventArgs e)
        {
            createNewBookmark(this.selected_node.Name);
        }

        private void goItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.selected_node.GetNodeCount(true) == 0 && this.selected_node.Name != rootNodeName)
                    Process.Start(GetHref(selected_node.Name));
            }
            catch
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, "Invalid URL");
            }
        }

        #endregion

        #region Generating thumbnail image

        Font font = new Font(FontFamily.GenericSerif, 16, FontStyle.Regular, GraphicsUnit.Pixel);
        PointF pointF;
        Brush brush = Brushes.Black;
        Pen pen = new Pen(Brushes.WhiteSmoke, 21);

        int previewImageWidth = 200;
        int previewImageHeight = 150;
        int previewImageBrowserWidth = 1024;
        int previewImageBrowserHeight = 768;

        #region pictureBox1

        private void generatePreviewItem1_Click(object sender, EventArgs e)
        {
            BackgroundWorker backgroundWorker1 = new BackgroundWorker();
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            //backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            //backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
            backgroundWorker1.RunWorkerAsync(this.selected_node.Text);
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            toolStripStatusLabel1.Text = "Creating Picturebox1.Image ...";

            try
            {

                WebsiteThumbnailImage thumbnailImage = new WebsiteThumbnailImage(this.GetHref(this.selected_node.Name), this.previewImageBrowserWidth, this.previewImageBrowserHeight, this.previewImageWidth, this.previewImageHeight);
                Image previewImage = thumbnailImage.GenerateWebSiteThumbnailImage();

                if (previewImage == null)
                {
                    toolStripStatusLabel1.Text = "Preview thumbnail image is null (Picturebox1)";
                    return;
                }

                this.pointF = new PointF(5, previewImage.Height - this.font.Size - 5);

                Graphics graphic1 = Graphics.FromImage(previewImage);
                graphic1.DrawRectangle(this.pen, 0, previewImage.Height - this.font.Size, previewImage.Width, previewImage.Height - (previewImage.Height - this.font.Size - 5));

                graphic1.DrawString((string)e.Argument, this.font, this.brush, this.pointF);
                this.pictureBox1.Image = previewImage;
                toolStripStatusLabel1.Text = "Preview thumbnail image created successfully (Picturebox1)";
            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = ex.Message;
            }
        }

        #endregion

        #region pictureBox2

        private void generatePreviewItem2_Click(object sender, EventArgs e)
        {
            BackgroundWorker backgroundWorker2 = new BackgroundWorker();
            backgroundWorker2.DoWork += new DoWorkEventHandler(backgroundWorker2_DoWork);
            backgroundWorker2.RunWorkerAsync(this.selected_node.Text);
        }

        private void backgroundWorker2_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                toolStripStatusLabel1.Text = "Creating Picturebox2.Image ...";

                WebsiteThumbnailImage thumbnailImage = new WebsiteThumbnailImage(this.GetHref(this.selected_node.Name), this.previewImageBrowserWidth, this.previewImageBrowserHeight, this.previewImageWidth, this.previewImageHeight);
                Image previewImage = thumbnailImage.GenerateWebSiteThumbnailImage();

                if (previewImage == null)
                {
                    toolStripStatusLabel1.Text = "Preview thumbnail image is null (Picturebox2)";
                    return;
                }

                this.pointF = new PointF(5, previewImage.Height - this.font.Size - 5);

                Graphics graphic1 = Graphics.FromImage(previewImage);
                graphic1.DrawRectangle(this.pen, 0, previewImage.Height - this.font.Size, previewImage.Width, previewImage.Height - (previewImage.Height - this.font.Size - 5));

                graphic1.DrawString((string)e.Argument, this.font, this.brush, this.pointF);
                this.pictureBox2.Image = previewImage;
                toolStripStatusLabel1.Text = "Preview thumbnail image created successfully (Picturebox2)";
            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = ex.Message;
            }
        }

        #endregion

        #region pictureBox3

        private void generatePreviewItem3_Click(object sender, EventArgs e)
        {
            BackgroundWorker backgroundWorker3 = new BackgroundWorker();
            backgroundWorker3.DoWork += new DoWorkEventHandler(backgroundWorker3_DoWork);
            backgroundWorker3.RunWorkerAsync(this.selected_node.Text);
        }

        private void backgroundWorker3_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                toolStripStatusLabel1.Text = "Creating Picturebox3.Image ...";

                WebsiteThumbnailImage thumbnailImage = new WebsiteThumbnailImage(this.GetHref(this.selected_node.Name), this.previewImageBrowserWidth, this.previewImageBrowserHeight, this.previewImageWidth, this.previewImageHeight);
                Image previewImage = thumbnailImage.GenerateWebSiteThumbnailImage();

                if (previewImage == null)
                {
                    toolStripStatusLabel1.Text = "Preview thumbnail image is null (Picturebox3)";
                    return;
                }

                this.pointF = new PointF(5, previewImage.Height - this.font.Size - 5);

                Graphics graphic1 = Graphics.FromImage(previewImage);
                graphic1.DrawRectangle(this.pen, 0, previewImage.Height - this.font.Size, previewImage.Width, previewImage.Height - (previewImage.Height - this.font.Size - 5));

                graphic1.DrawString((string)e.Argument, this.font, this.brush, this.pointF);
                this.pictureBox3.Image = previewImage;
                toolStripStatusLabel1.Text = "Preview thumbnail image created  successfully (Picturebox3)";
            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = ex.Message;
            }
        }


        #endregion

        #region pictureBox4

        private void generatePreviewItem4_Click(object sender, EventArgs e)
        {
            BackgroundWorker backgroundWorker4 = new BackgroundWorker();
            backgroundWorker4.DoWork += new DoWorkEventHandler(backgroundWorker4_DoWork);
            backgroundWorker4.RunWorkerAsync(this.selected_node.Text);
        }

        private void backgroundWorker4_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                toolStripStatusLabel1.Text = "Creating Picturebox4.Image ...";

                WebsiteThumbnailImage thumbnailImage = new WebsiteThumbnailImage(this.GetHref(this.selected_node.Name), this.previewImageBrowserWidth, this.previewImageBrowserHeight, this.previewImageWidth, this.previewImageHeight);
                Image previewImage = thumbnailImage.GenerateWebSiteThumbnailImage();

                if (previewImage == null)
                {
                    toolStripStatusLabel1.Text = "Preview thumbnail image is null (Picturebox4)";
                    return;
                }

                this.pointF = new PointF(5, previewImage.Height - this.font.Size - 5);

                Graphics graphic1 = Graphics.FromImage(previewImage);
                graphic1.DrawRectangle(this.pen, 0, previewImage.Height - this.font.Size, previewImage.Width, previewImage.Height - (previewImage.Height - this.font.Size - 5));

                graphic1.DrawString((string)e.Argument, this.font, this.brush, this.pointF);
                this.pictureBox4.Image = previewImage;
                toolStripStatusLabel1.Text = "Preview thumbnail image created successfully (Picturebox4)";

            }
            catch (Exception ex)
            {
                toolStripStatusLabel1.Text = ex.Message;
            }
        }

        #endregion

        #endregion

        #region contextMenuStripNotifyIcon

        void createContextMenuStripNotifyIcon()
        {
            try
            {
                if (this.treeViewBookMark.Nodes.Count <= 0) return;
                contextMenuStripNotifyIcon.Items.Clear();
                foreach (TreeNode node in this.treeViewBookMark.Nodes[0].Nodes)
                {
                    ToolStripMenuItem toolStripMenuItem = new ToolStripMenuItem();
                    if (node.GetNodeCount(true) > 0)
                    {
                        toolStripMenuItem.BackColor = contextMenuStripNotifyIcon.BackColor;
                        toolStripMenuItem.Image = imageList1.Images[node.ImageIndex];
                        toolStripMenuItem.Text = node.Text.Length > 30 ? node.Text.Substring(0, 30) + " ..." : node.Text;
                        toolStripMenuItem.Name = node.Name;
                        createContextMenuStripNotifyIcon(node, ref toolStripMenuItem);
                        contextMenuStripNotifyIcon.Items.Add(toolStripMenuItem);
                    }
                    else
                    {
                        toolStripMenuItem.BackColor = contextMenuStripNotifyIcon.BackColor;
                        toolStripMenuItem.Image = imageList1.Images[node.ImageIndex];
                        toolStripMenuItem.Text = node.Text.Length > 30 ? node.Text.Substring(0, 30) + " ..." : node.Text;
                        toolStripMenuItem.Name = node.Name;
                        toolStripMenuItem.MouseDown += new MouseEventHandler(contextMenuStripNotifyIconItems_MouseDown);
                        contextMenuStripNotifyIcon.Items.Add(toolStripMenuItem);
                    }
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        void createContextMenuStripNotifyIcon(TreeNode nodes, ref ToolStripMenuItem toolStripMenuItem)
        {
            foreach (TreeNode node in nodes.Nodes)
            {
                if (node.GetNodeCount(true) > 0)
                {
                    ToolStripMenuItem newToolStripMenuItem = new ToolStripMenuItem();
                    newToolStripMenuItem.BackgroundImage = imageList1.Images[node.ImageIndex];
                    newToolStripMenuItem.BackgroundImageLayout = ImageLayout.None;
                    newToolStripMenuItem.Text = node.Text.Length > 30 ? node.Text.Substring(0, 30) + " ..." : node.Text;
                    newToolStripMenuItem.Name = node.Name;
                    createContextMenuStripNotifyIcon(node, ref newToolStripMenuItem);
                    toolStripMenuItem.DropDown.Items.Add(newToolStripMenuItem);
                }
                else
                {
                    ToolStripMenuItem newToolStripMenuItem = new ToolStripMenuItem();
                    newToolStripMenuItem.BackgroundImage = imageList1.Images[node.ImageIndex];
                    newToolStripMenuItem.BackgroundImageLayout = ImageLayout.None;
                    newToolStripMenuItem.Text = node.Text.Length > 30 ? node.Text.Substring(0, 30) + " ..." : node.Text;
                    newToolStripMenuItem.Name = node.Name;
                    newToolStripMenuItem.MouseDown += new MouseEventHandler(contextMenuStripNotifyIconItems_MouseDown);
                    toolStripMenuItem.DropDown.Items.Add(newToolStripMenuItem);
                }
            }
        }

        void contextMenuStripNotifyIconItems_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Left)
                    Process.Start(GetHref(((ToolStripMenuItem)sender).Name));
            }
            catch
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, "Invalid URL");
            }
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region Notify Icon

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Visible = true;
            this.WindowState = FormWindowState.Normal;
            this.notifyIcon1.Visible = false;
            this.Activate();
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region menuStrip1 => File

        /// <summary>
        /// Shows About Me window
        /// </summary>
        private void aboutKingMarkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            about_me form = new about_me();
            form.ShowDialog();
        }

        private void exit(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure to closing this ?", "Warning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                Application.Exit();
        }

        private void optimize_menuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.XElementClopboard = null;
                this.TreeNodeClipboard = null;
                XElement optimizedXElement = new XElement(rootNodeName, new XAttribute("ELEMENTS", 0));

                int counter = 0;

                foreach (var element in this.X_Element.Elements())
                    optimizedXElement.Add(changeName(element, ref counter));

                optimizedXElement.Attribute("ELEMENTS").Value = counter.ToString();

                this.X_Element = optimizedXElement;
                XDocument xDoc = new XDocument(new XComment(XmlDescription), this.X_Element);
                xDoc.Save(this.temporaryFileName);

                listView1.Items.Clear();
                treeViewBookMark.Nodes.Clear();
                XmlToTree();

                successful("Finish", "XML file optimized successfully");
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        #endregion


        #region menuStrip1 => Import

        /// <summary>
        /// shows importing dialog and imports a Html file
        /// </summary>
        private void importHtml(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            List<string> list = new List<string>();
            openFileDialog1.Filter = "Html File(*.html;*.htm)|*.html;*.htm";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (StreamReader sr = new StreamReader(openFileDialog1.FileName))
                {
                    //Read lines
                    while (sr.Peek() >= 0)
                        list.Add(sr.ReadLine());
                }
                List<string> list2 = parse(list);
                if (traceToolStripMenuItem.Checked) trace(list2);
                ListToXml(list2);
                XmlToTree();
            }
        }

        /// <summary>
        /// shows importing dialog and imports a XML file
        /// </summary>
        private void importXML(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "XML File(*.xml)|*.xml";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    listView1.Items.Clear();
                    temporaryFileName = openFileDialog1.FileName;
                    X_Element = XElement.Load(temporaryFileName);
                    XmlToTree();
                }
                catch (Exception ex)
                {
                    StackFrame file_info = new StackFrame(true);
                    error(ref file_info, ex.Message);
                }
            }
        }


        #endregion


        #region menuStrip1 => Export to

        /// <summary>
        /// exports the KingMark XML file
        /// </summary>
        private void exportXML(object sender, EventArgs e)
        {
            update(selected_node.Name);

            try
            {
                saveFileDialog1.Filter = "XML File(*.xml)|*.xml";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    XDocument xDoc = new XDocument(new XComment(XmlDescription), this.X_Element);
                    xDoc.Save(saveFileDialog1.FileName);
                    successful("Finish", "All data saved successfully");
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        /// <summary>
        /// converts XML file to FF bookmark file
        /// I found an algoritm by experience ;) 
        /// </summary>
        private void exportHtml(object sender, EventArgs e)
        {
            update(selected_node.Name);//saves last changes
            createFFBookmark(this.X_Element);
        }

        void createFFBookmark(XElement xElement)
        {
            try
            {
                StringBuilder htmlContent = new StringBuilder("");//temporary buffer to save FF bookmark file in it
                StringBuilder spacer = new StringBuilder("    ");//number of space in beginning of each line in FF bookmark file
                try
                {
                    htmlContent.Append(@"<!DOCTYPE NETSCAPE-Bookmark-file-1>"
                                     + "\r\n<!-- This is an automatically generated file."
                                     + "\r\n     It will be read and overwritten."
                                     + "\r\n     DO NOT EDIT!"
                                     + "\r\n     Created by KingMark -->");
                    htmlContent.Append("\r\n<META HTTP-EQUIV=\"Content-Type\" CONTENT=\"text/html; charset=UTF-8\">");
                    htmlContent.Append("\r\n<TITLE>Bookmarks</TITLE>\r\n<H1>Bookmarks Menu</H1>");
                    htmlContent.Append("\r\n<DL><p>");

                    FileInfo tempFile = new FileInfo(Environment.GetEnvironmentVariable(TempVariable, EnvironmentVariableTarget.User) + "\\" + new FileInfo(this.temporaryFileName).Name);

                    this.X_Element.Save(this.temporaryFileName);
                    xElement.Save(tempFile.FullName);
                    string[] xmlLines = File.ReadAllLines(tempFile.FullName, Encoding.UTF8);

                    //Converts XML file to FF bookmark file                    
                    for (int i = 0; i < xmlLines.Length; i++)
                    {
                        if (Regex.IsMatch(xmlLines[i], "</NODE"))
                        {
                            if (spacer.Length > 4)
                                spacer.Length -= 4;
                            htmlContent.Append("\r\n" + spacer.ToString() + "</DL><p>");
                        }
                        else if (Regex.IsMatch(xmlLines[i], @"<NODE[^/]*/>"))
                        {
                            string name = Regex.Match(xmlLines[i], "NAME=\"[^\"]*\"").ToString();
                            name = Regex.Replace(name, "NAME=|\"", "");
                            name = System.Web.HttpUtility.HtmlDecode(name);

                            string dd = Regex.Match(xmlLines[i], "DESC=\"[^\"]*\"").ToString();
                            dd = Regex.Replace(dd, "DESC=|\"", "");
                            dd = System.Web.HttpUtility.HtmlDecode(dd);

                            htmlContent.Append(
                                "\r\n" + spacer.ToString()
                                + "<DT><H3 " + System.Web.HttpUtility.HtmlDecode(Regex.Match(xmlLines[i], "ADD_DATE=\"[^\"]*\"").ToString())
                                + " " + System.Web.HttpUtility.HtmlDecode(Regex.Match(xmlLines[i], "LAST_MODIFIED=\"[^\"]*\"").ToString())
                                + ">"
                                + name
                                + "</H3>\r\n"
                                + spacer.ToString() + "<DL><p>");
                            if (dd != "") htmlContent.Append("\r\n<DD>" + dd);
                            htmlContent.Append("\r\n" + spacer.ToString() + "</DL><p>");
                        }
                        else if (Regex.IsMatch(xmlLines[i], @"<NODE"))
                        {
                            if (Regex.IsMatch(xmlLines[i - 1], @"<NODE")) spacer.Append("    ");

                            string name = Regex.Match(xmlLines[i], "NAME=\"[^\"]*\"").ToString();
                            name = Regex.Replace(name, "NAME=|\"", "");
                            name = System.Web.HttpUtility.HtmlDecode(name);

                            string dd = Regex.Match(xmlLines[i], "DESC=\"[^\"]*\"").ToString();
                            dd = Regex.Replace(dd, "DESC=|\"", "");
                            dd = System.Web.HttpUtility.HtmlDecode(dd);

                            htmlContent.Append(
                                "\r\n" + spacer.ToString()
                                + "<DT><H3 " + System.Web.HttpUtility.HtmlDecode(Regex.Match(xmlLines[i], "ADD_DATE=\"[^\"]*\"").ToString())
                                + " " + System.Web.HttpUtility.HtmlDecode(Regex.Match(xmlLines[i], "LAST_MODIFIED=\"[^\"]*\"").ToString())
                                + ">"
                                + name
                                + "</H3>\r\n"
                                + spacer.ToString() + "<DL><p>");
                            if (dd != "") htmlContent.Append("\r\n<DD>" + dd);
                        }
                        else if (Regex.IsMatch(xmlLines[i], @"\<BOOKMARK"))
                        {
                            if (i > 0)
                                if (Regex.IsMatch(xmlLines[i - 1], @"<NODE")) spacer.Append("    ");

                            string dd = Regex.Match(xmlLines[i], "DESC=\"[^\"]*\"").ToString();
                            dd = Regex.Replace(dd, "DESC=|\"", "");
                            dd = System.Web.HttpUtility.HtmlDecode(dd);

                            string NAME = Regex.Match(xmlLines[i], "NAME=\"[^\"]*\"").ToString();
                            NAME = Regex.Replace(NAME, "NAME=|\"", "");
                            NAME = System.Web.HttpUtility.HtmlDecode(NAME);

                            htmlContent.Append(
                                "\r\n" + spacer.ToString()
                                + "<DT><A " + System.Web.HttpUtility.HtmlDecode(Regex.Match(xmlLines[i], "HREF=\"[^\"]*\"").ToString())
                                + " " + System.Web.HttpUtility.HtmlDecode(Regex.Match(xmlLines[i], "ADD_DATE=\"[^\"]*\"").ToString())
                                + " " + System.Web.HttpUtility.HtmlDecode(Regex.Match(xmlLines[i], "LAST_MODIFIED=\"[^\"]*\"").ToString())
                                + " " + System.Web.HttpUtility.HtmlDecode(Regex.Match(xmlLines[i], "ICON_URI=\"[^\"]*\"").ToString())
                                + " " + System.Web.HttpUtility.HtmlDecode(Regex.Match(xmlLines[i], "ICON=\"[^\"]*\"").ToString())
                                + " " + System.Web.HttpUtility.HtmlDecode(Regex.Match(xmlLines[i], "SHORTCUTURL=\"[^\"]*\"").ToString())
                                + " LAST_CHARSET=\"UTF-8\""
                                + ">"
                                + NAME
                                + "</A>");
                            if (dd != "") htmlContent.Append("\r\n<DD>" + dd);
                        }
                    }

                    if (spacer.Length > 0) spacer.Length -= 4;
                    htmlContent.Append("\r\n" + spacer.ToString() + "</DL><p>");

                    tempFile.Delete();
                }
                catch
                {
                    return;
                }

                //shows saveFileDialog & saves file
                try
                {
                    saveFileDialog1.Filter = "Html File(*.html;*.htm)|*.html;*.htm";
                    saveFileDialog1.FileName = "KingMark.html";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        StreamWriter sw = new StreamWriter(saveFileDialog1.FileName);

                        //Write to the text file
                        sw.Write(htmlContent.ToString());
                        //Close the file
                        sw.Close();
                        successful("Finish", "All data saved successfully");
                    }
                }
                catch (Exception ex)
                {
                    StackFrame file_info = new StackFrame(true);
                    error(ref file_info, ex.Message);
                }

                spacer.Length = 0;
                htmlContent.Length = 0;
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        /// <summary>
        /// For navigation folders target to creating IE favorites
        /// </summary>
        DirectoryInfo directoryInfo;

        private void exportIE_favorites_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                directoryInfo = new DirectoryInfo(folderBrowserDialog1.SelectedPath + "//" + rootNodeName);
                Directory.CreateDirectory(directoryInfo.FullName);

                foreach (var element in X_Element.Elements())
                    exportIE_favorite(element.Name.ToString());
                successful("Finish", "All data saved successfully");
            }
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region TreeView

        /// <summary>
        /// Renames node's label and save it in the temporary XML file
        /// </summary>
        private void treeViewBookMark_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            try
            {
                selected_node = e.Node;
                if (selected_node.Name != rootNodeName)
                {
                    if (e.Label == null || e.Label.Trim() == "")
                    {
                        e.CancelEdit = true;
                        return;
                    }

                    var SelectedNode = (from node in X_Element.Descendants(e.Node.Name)
                                        select node).First();

                    SelectedNode.SetAttributeValue("NAME", e.Label.Trim());
                    X_Element.Save(this.temporaryFileName);

                    treeViewBookMark.SelectedNode = selected_node;
                    textBoxName.Text = e.Label.Trim();
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                e.CancelEdit = true;
            }
        }

        private void treeViewBookMark_AfterSelect(object sender, TreeViewEventArgs e)
        {
            this.selected_node = e.Node;
            string nodeName = selected_node.Name;

            //by this calendar (MDCalendar) I converted ADD_DATE & LAST_MODIFIED to the correct date
            //Learning more : http://www.codeproject.com/KB/cs/PersianCalendar.aspx
            MDCalendar calendar = new MDCalendar();
            try
            {
                if (selected_node.Name == rootNodeName)
                {
                    textBoxLastModified.Text = textBoxAddDate.Text = textBoxIconURL.Text = textBoxKeyWords.Text = textBoxKeyWords.Text = textBoxHref.Text = textBoxName.Text = "";

                    textBoxIconURL.Enabled = textBoxDesc.Enabled = textBoxName.Enabled = buttonReloadIcon.Enabled = buttonDelIcon.Enabled = textBoxKeyWords.Enabled = textBoxHref.Enabled = false;
                    generatePreviewItem.Enabled = deleteItem.Enabled = cutItem.Enabled = copyItem.Enabled = goItem.Enabled = false;
                    pasteItem.Enabled = this.XElementClopboard == null ? false : true;
                }
                else if (Regex.IsMatch(this.selected_node.Name, @"NODE"))
                {
                    textBoxLastModified.Text = textBoxAddDate.Text = textBoxIconURL.Text = textBoxKeyWords.Text = textBoxKeyWords.Text = textBoxHref.Text = "";

                    buttonReloadIcon.Enabled = buttonDelIcon.Enabled = textBoxKeyWords.Enabled = textBoxHref.Enabled = textBoxIconURL.Enabled = false;
                    textBoxDesc.Enabled = textBoxName.Enabled = true;

                    deleteItem.Enabled = cutItem.Enabled = copyItem.Enabled = true;
                    generatePreviewItem.Enabled = goItem.Enabled = false;
                    pasteItem.Enabled = this.XElementClopboard == null ? false : true;

                    pictureBoxIcon.Image = null;
                    textBoxName.Text = selected_node.Text;
                    textBoxDesc.Text = GetDescription(nodeName);
                    textBoxAddDate.Text = (GetAddDate(nodeName) == 0) ? "" : calendar.Date("Z/e/d  h:s A ", GetAddDate(nodeName));
                    textBoxLastModified.Text = (GetLastModified(nodeName) == 0) ? "" : calendar.Date("Z/e/d  h:s A ", GetLastModified(nodeName));
                }
                else if (Regex.IsMatch(this.selected_node.Name, @"BOOKMARK"))
                {
                    buttonReloadIcon.Enabled = textBoxIconURL.Enabled = textBoxDesc.Enabled = textBoxName.Enabled = buttonDelIcon.Enabled = textBoxKeyWords.Enabled = textBoxHref.Enabled = true;
                    generatePreviewItem.Enabled = deleteItem.Enabled = cutItem.Enabled = copyItem.Enabled = true;
                    deleteItem.Enabled = cutItem.Enabled = copyItem.Enabled = goItem.Enabled = true;
                    pasteItem.Enabled = this.XElementClopboard == null ? false : true;

                    pictureBoxIcon.Image = createIcon(selected_node.Name);
                    textBoxName.Text = e.Node.Text;
                    textBoxHref.Text = GetHref(nodeName);
                    textBoxDesc.Text = GetDescription(nodeName);
                    textBoxKeyWords.Text = GetKeyword(nodeName);
                    textBoxIconURL.Text = GetIconUri(nodeName);
                    textBoxAddDate.Text = (GetAddDate(nodeName) == 0) ? "" : calendar.Date("Z/e/d  h:s A ", GetAddDate(nodeName));
                    textBoxLastModified.Text = (GetLastModified(nodeName) == 0) ? "" : calendar.Date("Z/e/d  h:s A ", GetLastModified(nodeName));
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        /// <summary>
        /// Manages DragDrop operations
        /// </summary>   
        private void treeViewBookMark_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                #region Firefox tab
                if (e.Data.GetDataPresent(DataFormats.StringFormat))
                {
                    string location = e.Data.GetData("Text").ToString();

                    if (Regex.IsMatch(location, @"http:\/\/(\w\w\w\.)?[^\/]*(\/)?"))
                    {
                        int max_elements = Int32.Parse(X_Element.Attribute("ELEMENTS").Value);
                        max_elements++;
                        MDCalendar calendar = new MDCalendar();
                        DateTime date = DateTime.Now;
                        TimeZone time = TimeZone.CurrentTimeZone;
                        TimeSpan difference = time.GetUtcOffset(date);
                        uint currentTime = calendar.Time() + (uint)difference.TotalSeconds;

                        XElement newXmlNode = new XElement("BOOKMARK" + max_elements.ToString(),
                                                        new XAttribute("NAME", location),
                                                        new XAttribute("HREF", location),
                                                        new XAttribute("SHORTCUTURL", ""),
                                                        new XAttribute("DESC", ""),
                                                        new XAttribute("ADD_DATE", currentTime.ToString()),
                                                        new XAttribute("LAST_MODIFIED", currentTime.ToString()),
                                                        new XAttribute("ICON_URI", ""),
                                                        new XAttribute("ICON", ""));

                        Point pt = ((TreeView)sender).PointToClient(new Point(e.X, e.Y));
                        TreeNode destinationTreeNode = ((TreeView)sender).GetNodeAt(pt);

                        TreeNode newTreeNode = new TreeNode(newXmlNode.Attribute("NAME").Value);
                        newTreeNode.Name = newXmlNode.Name.ToString();
                        newTreeNode.ImageIndex = 1;
                        newTreeNode.SelectedImageIndex = 1;

                        //if DestinationNode equels null I inserted the new node after the last node
                        if (destinationTreeNode == null || Regex.IsMatch(destinationTreeNode.Name, rootNodeName))
                        {
                            TreeView treeView = ((TreeView)sender);
                            int lastNode_index = treeView.Nodes[0].Nodes.Count - 1;

                            if (lastNode_index < 0)
                            {
                                destinationTreeNode = treeView.Nodes[0];
                                this.X_Element.Add(newXmlNode);
                            }
                            else
                            {
                                destinationTreeNode = treeView.Nodes[0].Nodes[lastNode_index];//selects latest node
                                this.X_Element.Element(destinationTreeNode.Name).AddAfterSelf(newXmlNode);
                            }

                            destinationTreeNode = treeView.Nodes[0];//selects latest node for TreeView
                            destinationTreeNode.Nodes.Add((TreeNode)newTreeNode.Clone());
                        }
                        else if (Regex.IsMatch(destinationTreeNode.Name, "NODE"))
                        {
                            XElement destinationElement = (from element in this.X_Element.Descendants(destinationTreeNode.Name)
                                                           select element).First();
                            destinationElement.Add(newXmlNode);
                            destinationTreeNode.Nodes.Add(newTreeNode);
                        }
                        else if (Regex.IsMatch(destinationTreeNode.Name, "BOOKMARK"))
                        {
                            XElement destinationElement = (from element in this.X_Element.Descendants(destinationTreeNode.Name)
                                                           select element).First().Parent;
                            destinationElement.Add(newXmlNode);
                            destinationTreeNode.Parent.Nodes.Add(newTreeNode);
                        }

                        this.treeViewBookMark.SelectedNode = this.selected_node = this.treeViewBookMark.Nodes.Find("BOOKMARK" + max_elements.ToString(), true).First();
                        this.treeViewBookMark.SelectedNode.Expand();
                        X_Element.Attribute("ELEMENTS").Value = max_elements.ToString();
                    }
                }
                #endregion

                #region IE bookmarks *.url

                else if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    foreach (var file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);
                        if (!Regex.IsMatch(fileInfo.Extension, "url", RegexOptions.IgnoreCase)) return;
                        string location = "";
                        string iconUru = "";
                        using (StreamReader sr = new StreamReader(file))
                        {
                            string line = "";
                            //Read lines
                            while (sr.Peek() >= 0)
                            {
                                line = sr.ReadLine();
                                if (Regex.IsMatch(line, @"URL=", RegexOptions.IgnoreCase))
                                    location = Regex.Replace(line, @"URL=", "", RegexOptions.IgnoreCase).Trim();
                                else if (Regex.IsMatch(line, @"IconFile=", RegexOptions.IgnoreCase))
                                    iconUru = Regex.Replace(line, @"IconFile=", "", RegexOptions.IgnoreCase).Trim();
                            }
                        }

                        int max_elements = Int32.Parse(X_Element.Attribute("ELEMENTS").Value);
                        max_elements++;
                        MDCalendar calendar = new MDCalendar();
                        DateTime date = DateTime.Now;
                        TimeZone time = TimeZone.CurrentTimeZone;
                        TimeSpan difference = time.GetUtcOffset(date);
                        uint currentTime = calendar.Time() + (uint)difference.TotalSeconds;

                        XElement newXmlNode = new XElement("BOOKMARK" + max_elements.ToString(),
                                                        new XAttribute("NAME", Regex.Replace(fileInfo.Name, fileInfo.Extension, "")),
                                                        new XAttribute("HREF", location),
                                                        new XAttribute("SHORTCUTURL", ""),
                                                        new XAttribute("DESC", ""),
                                                        new XAttribute("ADD_DATE", currentTime.ToString()),
                                                        new XAttribute("LAST_MODIFIED", currentTime.ToString()),
                                                        new XAttribute("ICON_URI", iconUru),
                                                        new XAttribute("ICON", ""));

                        Point pt = ((TreeView)sender).PointToClient(new Point(e.X, e.Y));
                        TreeNode destinationTreeNode = ((TreeView)sender).GetNodeAt(pt);

                        TreeNode newTreeNode = new TreeNode(newXmlNode.Attribute("NAME").Value);
                        newTreeNode.Name = newXmlNode.Name.ToString();
                        newTreeNode.ImageIndex = 1;
                        newTreeNode.SelectedImageIndex = 1;

                        //if DestinationNode equels null I inserted the new node after the latest node
                        if (destinationTreeNode == null || Regex.IsMatch(destinationTreeNode.Name, rootNodeName))
                        {
                            TreeView treeView = ((TreeView)sender);
                            int lastNode_index = treeView.Nodes[0].Nodes.Count - 1;

                            if (lastNode_index < 0)
                            {
                                destinationTreeNode = treeView.Nodes[0];
                                this.X_Element.Add(newXmlNode);
                            }
                            else
                            {
                                destinationTreeNode = treeView.Nodes[0].Nodes[lastNode_index];//selects latest node
                                this.X_Element.Element(destinationTreeNode.Name).AddAfterSelf(newXmlNode);
                            }
                            destinationTreeNode = treeView.Nodes[0];//selects latest node for TreeView
                            destinationTreeNode.Nodes.Add((TreeNode)newTreeNode.Clone());
                        }
                        else if (Regex.IsMatch(destinationTreeNode.Name, "NODE"))
                        {
                            XElement destinationElement = (from element in this.X_Element.Descendants(destinationTreeNode.Name)
                                                           select element).First();
                            destinationElement.Add(newXmlNode);
                            destinationTreeNode.Nodes.Add(newTreeNode);
                        }
                        else if (Regex.IsMatch(destinationTreeNode.Name, "BOOKMARK"))
                        {
                            XElement destinationElement = (from element in this.X_Element.Descendants(destinationTreeNode.Name)
                                                           select element).First().Parent;
                            destinationElement.Add(newXmlNode);
                            destinationTreeNode.Parent.Nodes.Add(newTreeNode);
                        }

                        this.treeViewBookMark.SelectedNode = this.selected_node = this.treeViewBookMark.Nodes.Find("BOOKMARK" + max_elements.ToString(), true).First();
                        this.treeViewBookMark.SelectedNode.Expand();
                        X_Element.Attribute("ELEMENTS").Value = max_elements.ToString();
                    }
                }

                #endregion

                this.X_Element.Save(temporaryFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void treeViewBookMark_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void treeViewBookMark_DragLeave(object sender, EventArgs e)
        {
            try
            {
                if (this.selected_node.GetNodeCount(true) == 0)
                {
                    string location = GetHref(selected_node.Name);

                    string file_name = this.selected_node.Text;
                    file_name = Regex.Replace(HttpUtility.HtmlDecode(file_name), @"http:\/\/|ftp:\/\/", "");
                    file_name = Regex.Replace(file_name, @"\/|//|:|\*|\?|\<|\>|\|" + "|\"", "");
                    file_name = file_name.Length > 250 ? file_name.Substring(0, 250) : file_name;

                    string[] files = new string[] { Environment.GetEnvironmentVariable(TempVariable, EnvironmentVariableTarget.User) + "//" + file_name + ".url" };                    

                    foreach (var file in files)
                    {
                        if (File.Exists(file)) File.Delete(file);
                        StreamWriter sw = new StreamWriter(file, true);
                        sw.Write("[InternetShortcut]" + Environment.NewLine + "URL=" + location);
                        sw.Close();
                    }

                    //Create a new data object with the FileDrop type                    
                    DataObject data = new DataObject(DataFormats.FileDrop, files);
                    this.treeViewBookMark.DoDragDrop(data, DragDropEffects.Move);
                }
                else
                {
                    string rootFolder = this.selected_node.Text;
                    rootFolder = Regex.Replace(HttpUtility.HtmlDecode(rootFolder), @"http:\/\/|ftp:\/\/", "");
                    rootFolder = Regex.Replace(rootFolder, @"\/|//|:|\*|\?|\<|\>|\|" + "|\"", "");
                    rootFolder = rootFolder.Length > 250 ? rootFolder.Substring(0, 250) : rootFolder;

                    directoryInfo = new DirectoryInfo(Environment.GetEnvironmentVariable(TempVariable, EnvironmentVariableTarget.User) + "\\" + rootFolder);
                    Directory.CreateDirectory(directoryInfo.FullName);

                    string[] files = new string[] { directoryInfo.FullName };

                    if (this.selected_node.Name == rootNodeName)
                    {
                        foreach (var element in this.X_Element.Elements())
                            exportIE_favorite(element.Name.ToString());
                    }
                    else
                    {
                        var xElement = (from n in this.X_Element.Descendants(this.selected_node.Name) select n).Elements();
                        foreach (var element in xElement)
                            exportIE_favorite(element.Name.ToString());
                    }

                    //Create a new data object with the FileDrop type
                    DataObject data = new DataObject(DataFormats.FileDrop, files);
                    this.treeViewBookMark.DoDragDrop(data, DragDropEffects.Move);
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void treeViewBookMark_ItemDrag(object sender, ItemDragEventArgs e)
        {
            selected_node = (TreeNode)e.Item;
            this.treeViewBookMark.DoDragDrop(selected_node, DragDropEffects.Copy);
        }

        private void treeViewBookMark_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                selected_node = treeViewBookMark.SelectedNode;

                if (e.KeyCode == Keys.Delete)
                {
                    if (MessageBox.Show("Are you sure ?", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                        deleteNodes(selected_node.Name);
                }
                else if (e.KeyCode == Keys.Enter)
                {
                    if (selected_node.GetNodeCount(true) == 0 && selected_node.Name != rootNodeName)
                        Process.Start(GetHref(selected_node.Name));
                }
                else if (e.Control && e.KeyCode == Keys.C)
                {
                    copyToClipboard(selected_node.Name);
                }
                else if (e.Control && e.KeyCode == Keys.X)
                {
                    cutToClipboard(selected_node.Name);
                }
                else if (e.Control && e.KeyCode == Keys.V)
                {
                    pasteClipboard();
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void treeViewBookMark_MouseClick(object sender, MouseEventArgs e)
        {
            selected_node = treeViewBookMark.GetNodeAt(new Point(e.X, e.Y));
            if (e.Button == MouseButtons.Right)
            {
                treeViewBookMark.SelectedNode = selected_node;
                contextMenuStripTreeView.Show(this.treeViewBookMark, new Point(e.X, e.Y));
            }
        }

        /// <summary>
        /// Opens Internet browser  
        /// </summary>
        private void treeViewBookMark_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            try
            {
                selected_node = e.Node;
                if (selected_node.GetNodeCount(true) == 0 && selected_node.Name != rootNodeName)
                    Process.Start(GetHref(e.Node.Name));
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region TextBoxes

        private void textBox_Validated(object sender, EventArgs e)
        {
            try
            {
                update(selected_node.Name);
            }
            catch
            { }
        }

        private void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                    update(selected_node.Name);
                else if (e.Control && e.KeyCode == Keys.A)
                    ((TextBox)sender).SelectAll();
            }
            catch
            { }
        }

        /// <summary>
        /// Searchs XML file for finding a needful node
        /// </summary>
        private void textBoxSearchTxt_TextChanged(object sender, EventArgs e)
        {
            string searchTxt = Regex.Replace(textBoxSearchTxt.Text.Trim(), @"[\$\^\{\[\(\|\)\*\+\?\\]", "");
            if (searchTxt == "")
            {
                this.listView1.Items.Clear();
                return;
            }

            try
            {
                var searchNodes = from nodes in X_Element.Descendants()
                                  where
                                  (nodes.Attribute("NAME") != null && Regex.IsMatch(nodes.Attribute("NAME").Value, searchTxt, RegexOptions.IgnoreCase))
                                   || (nodes.Attribute("HREF") != null && Regex.IsMatch(nodes.Attribute("HREF").Value, searchTxt, RegexOptions.IgnoreCase))
                                   || (nodes.Attribute("SHORTCUTURL") != null && Regex.IsMatch(nodes.Attribute("SHORTCUTURL").Value, searchTxt, RegexOptions.IgnoreCase))
                                  select nodes;

                listView1.Items.Clear();

                foreach (var node in searchNodes)
                {
                    try
                    {
                        ListViewItem item = new ListViewItem(new string[] { node.Attribute("NAME").Value, node.Attribute("HREF").Value });
                        item.Name = node.Name.ToString();
                        listView1.Items.Add(item);
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch { }
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region Buttons

        #region Delete icons

        private void deleteIcon_Click(object sender, EventArgs e)
        {
            selected_node = treeViewBookMark.SelectedNode;
            clearIcon(selected_node.Name);
            pictureBoxIcon.Image = null;
            selected_node.ImageIndex = 1;
            selected_node.SelectedImageIndex = 1;
        }

        private void buttonDelIcon_MouseLeave(object sender, EventArgs e)
        {
            toolTip1.Hide(this.buttonDelIcon);
        }

        private void buttonDelIcon_MouseMove(object sender, MouseEventArgs e)
        {
            toolTip1.Show("Delete this icon", this.buttonDelIcon, e.X + 10, e.Y + 10);
        }

        #endregion

        #region Reload icons

        private void buttonReloadIcon_Click(object sender, EventArgs e)
        {
            this.selected_node = treeViewBookMark.SelectedNode;
            reloadIcon(selected_node.Name);
        }

        private void buttonReloadIcon_MouseLeave(object sender, EventArgs e)
        {
            toolTip1.Hide(this.buttonDelIcon);
        }

        private void buttonReloadIcon_MouseMove(object sender, MouseEventArgs e)
        {
            toolTip1.Show("Reload the icon", this.buttonReloadIcon, e.X + 10, e.Y + 10);
        }

        #endregion

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region listView

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            treeViewBookMark.CollapseAll();
            try
            {
                this.treeViewBookMark.SelectedNode = this.selected_node =
                    this.treeViewBookMark.Nodes.Find(((ListView)sender).SelectedItems[0].Name, true).First();
            }
            catch { }
        }

        #endregion
    }
}