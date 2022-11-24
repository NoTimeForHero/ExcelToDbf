using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ExcelToDbf.Core.ViewModels;

namespace ExcelToDbf.Core.Views
{
    /// <summary>
    /// Логика взаимодействия для EditPreloadView.xaml
    /// </summary>
    public partial class EditPreloadView : Window
    {
        public EditPreloadView(EditPreloadVM context)
        {
            DataContext = context;
            InitializeComponent();
            btnClose.Click += (o, ev) => Close();

            Closed += EditPreloadView_Closed;
        }

        private void EditPreloadView_Closed(object sender, EventArgs e)
        {
            var model = (EditPreloadVM)DataContext;
            if (model == null) return;
            model.ReloadCommand.Execute(new System.Reactive.Unit());
        }
    }
}
