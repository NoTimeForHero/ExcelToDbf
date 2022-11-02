using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ReactiveUI;
using Unity;

namespace ExcelToDbf.Utils.Extensions
{
    public static class UnityExtensions
    {
        public static IUnityContainer RegisterSingletonMVVM<TView, TViewModel>(this IUnityContainer container)
            where TView : FrameworkElement
            where TViewModel : ReactiveObject
        {
            container.RegisterSingleton<TView>();
            container.RegisterSingleton<TViewModel>();
            return container;
        }
    }
}
