using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace Desktop_Browser
{
    internal class ClickEventAdder
    {
        public event RoutedEventHandler Click;

        public ClickEventAdder(UIElement element) {
            element.PreviewMouseLeftButtonUp += (sender, e) => Click(element, new RoutedEventArgs());
            element.PreviewTouchUp += (sender, e) => Click(element, new RoutedEventArgs());
            element.PreviewStylusUp += (sender, e) => Click(element, new RoutedEventArgs());
            element.Focusable = true;
            element.PreviewKeyUp += (object sender,KeyEventArgs e) => { if (e.Key == Key.Enter) { Click(element, new RoutedEventArgs()); } };
        } 
    }
}
