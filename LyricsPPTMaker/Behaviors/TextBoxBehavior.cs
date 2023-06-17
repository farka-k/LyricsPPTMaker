using Microsoft.Xaml.Behaviors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace LyricsPPTMaker
{
    public class TextBoxBehavior:Behavior<TextBox>
    {
        public bool IsSelectAll { get; set; } = true;
        public static readonly DependencyProperty IsControlFocusProperty =
            DependencyProperty.Register(nameof(IsControlFocus), typeof(bool), typeof(TextBoxBehavior), new PropertyMetadata(false, propertyChangedCallback: IsControlFocusPropertyChanged));

        public bool IsControlFocus
        {
            get { return (bool)GetValue(IsControlFocusProperty); }
            set { SetValue(IsControlFocusProperty, value); }
        }

        private static void IsControlFocusPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TextBoxBehavior behavior)
            {
                behavior.AssociatedObject.Focus();
                behavior.AssociatedObject.SelectAll();
            }
        }
        protected override void OnAttached()
        {
            if (IsSelectAll)
            {
                AssociatedObject.GotFocus += AssociatedObject_GotFocus;
            }
            AssociatedObject.KeyDown += AssociatedObject_KeyDown;

        }

        private void AssociatedObject_KeyDown(object sender, KeyEventArgs e)
        {
            //Enter Key입력시 EnterCommand 실행 이벤트 아규먼트 넘기는 것은 옵션
            if (e.Key == Key.Enter && EnterCommand != null)
            {
                EnterCommand.Execute(e);
                e.Handled = true;
            }
        }

        private void AssociatedObject_GotFocus(object sender, System.Windows.RoutedEventArgs e)
        {
            //Dispatcher를 이용해야 기능이 정상 동작합니다.
            Dispatcher.BeginInvoke(
                () =>
                {
                    AssociatedObject.SelectAll();
                }, null);
        }

        protected override void OnDetaching()
        {
            if (IsSelectAll)
            {
                AssociatedObject.GotFocus -= AssociatedObject_GotFocus;
            }
            AssociatedObject.KeyDown -= AssociatedObject_KeyDown;
        }

        public ICommand EnterCommand
        {
            get { return (ICommand)GetValue(EnterCommandProperty); }
            set { SetValue(EnterCommandProperty, value); }
        }

        /// <summary>
        /// EnterCommand
        /// </summary>
        public static readonly DependencyProperty EnterCommandProperty =
            DependencyProperty.Register(nameof(EnterCommand), typeof(ICommand), typeof(TextBoxBehavior), new PropertyMetadata(null));
    }
}
