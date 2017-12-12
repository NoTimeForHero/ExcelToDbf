using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using ExcelToDbf.Properties;

namespace ExcelToDbf.Sources.View
{
    partial class AboutBox : Form
    {
        public AboutBox()
        {
            InitializeComponent();
            Text = $"О программе: {AssemblyTitle}";
            labelProductName.Text += $": {AssemblyProduct}";
            labelVersion.Text = $"Версия: {AssemblyVersion}";
            labelCompanyName.Text += $": {AssemblyCompany}";

            string about = "Excel® является зарегистрированной торговой маркой Microsoft." + Environment.NewLine;
            about += $"Разработчик программы: <a href='http://github.com/{AssemblyCompany}'>{AssemblyCompany}</a> <br/>";
            about += "Все права на иконки принадлежат их авторам: <br/>";

            Type resourceType = typeof(IconCredits);
            PropertyInfo[] resourceProps = resourceType.GetProperties( BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.GetProperty);

            int count = 0;
            foreach (PropertyInfo info in resourceProps)
            {
                if (info.PropertyType != typeof(string)) continue;

                string value = info.GetValue(null, null) as string;
                if (value == null) break;

                string[] parts = value.Split(new char[]{';'}, 2);
                about += $"<a href='{parts[1]}'>{parts[0]}</a>, ";
                count++;
            }
            webBrowser1.DocumentText = about;
        }

        public sealed override string Text
        {
            get { return base.Text; }
            set { base.Text = value; }
        }

        #region Методы доступа к атрибутам сборки

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion

        private void okButton_Click(object sender, EventArgs e)
        {
            Close();
        } 

        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            if (e.Url == new Uri("about:blank")) return;
            e.Cancel = true;
            System.Diagnostics.Process.Start(e.Url.AbsoluteUri);
        }

        private void AboutBox_Load(object sender, EventArgs e)
        {

        }
    }
}
