﻿#pragma checksum "..\..\MoveMoneyWindow.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "70B3BC5096450A3447DC217E297625D9"
//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан программой.
//     Исполняемая версия:4.0.30319.34209
//
//     Изменения в этом файле могут привести к неправильной работе и будут потеряны в случае
//     повторной генерации кода.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;


namespace CreditApp {
    
    
    /// <summary>
    /// MoveMoneyWindow
    /// </summary>
    public partial class MoveMoneyWindow : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 19 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox MoveComboBox;
        
        #line default
        #line hidden
        
        
        #line 30 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label DocumentNameLabel;
        
        #line default
        #line hidden
        
        
        #line 31 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox DocNamberTexBox;
        
        #line default
        #line hidden
        
        
        #line 44 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.DatePicker DatePicker;
        
        #line default
        #line hidden
        
        
        #line 59 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label SummTextBox;
        
        #line default
        #line hidden
        
        
        #line 66 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox MoveMoneyComboBox;
        
        #line default
        #line hidden
        
        
        #line 78 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox DebitMoneyTextBox;
        
        #line default
        #line hidden
        
        
        #line 83 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox OsnovanieTextBox;
        
        #line default
        #line hidden
        
        
        #line 98 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label LocalSumm;
        
        #line default
        #line hidden
        
        
        #line 101 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button AddButton;
        
        #line default
        #line hidden
        
        
        #line 106 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button SaveButton;
        
        #line default
        #line hidden
        
        
        #line 120 "..\..\MoveMoneyWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.DataGrid DataGrid;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/CreditApp;component/movemoneywindow.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\MoveMoneyWindow.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.MoveComboBox = ((System.Windows.Controls.ComboBox)(target));
            
            #line 19 "..\..\MoveMoneyWindow.xaml"
            this.MoveComboBox.SelectionChanged += new System.Windows.Controls.SelectionChangedEventHandler(this.MoveComboBox_SelectionChanged);
            
            #line default
            #line hidden
            return;
            case 2:
            this.DocumentNameLabel = ((System.Windows.Controls.Label)(target));
            return;
            case 3:
            this.DocNamberTexBox = ((System.Windows.Controls.TextBox)(target));
            
            #line 34 "..\..\MoveMoneyWindow.xaml"
            this.DocNamberTexBox.TextChanged += new System.Windows.Controls.TextChangedEventHandler(this.DocNamberTexBox_TextChanged);
            
            #line default
            #line hidden
            return;
            case 4:
            this.DatePicker = ((System.Windows.Controls.DatePicker)(target));
            return;
            case 5:
            this.SummTextBox = ((System.Windows.Controls.Label)(target));
            return;
            case 6:
            this.MoveMoneyComboBox = ((System.Windows.Controls.ComboBox)(target));
            
            #line 70 "..\..\MoveMoneyWindow.xaml"
            this.MoveMoneyComboBox.SelectionChanged += new System.Windows.Controls.SelectionChangedEventHandler(this.MoveMoneyComboBox_SelectionChanged);
            
            #line default
            #line hidden
            return;
            case 7:
            this.DebitMoneyTextBox = ((System.Windows.Controls.TextBox)(target));
            
            #line 78 "..\..\MoveMoneyWindow.xaml"
            this.DebitMoneyTextBox.TextChanged += new System.Windows.Controls.TextChangedEventHandler(this.DebitMoneyTextBox_TextChanged);
            
            #line default
            #line hidden
            return;
            case 8:
            this.OsnovanieTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 9:
            this.LocalSumm = ((System.Windows.Controls.Label)(target));
            return;
            case 10:
            this.AddButton = ((System.Windows.Controls.Button)(target));
            
            #line 102 "..\..\MoveMoneyWindow.xaml"
            this.AddButton.Click += new System.Windows.RoutedEventHandler(this.AddButton_Click);
            
            #line default
            #line hidden
            return;
            case 11:
            this.SaveButton = ((System.Windows.Controls.Button)(target));
            
            #line 107 "..\..\MoveMoneyWindow.xaml"
            this.SaveButton.Click += new System.Windows.RoutedEventHandler(this.SaveButton_Click);
            
            #line default
            #line hidden
            return;
            case 12:
            
            #line 112 "..\..\MoveMoneyWindow.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.Button_Click_Exit);
            
            #line default
            #line hidden
            return;
            case 13:
            this.DataGrid = ((System.Windows.Controls.DataGrid)(target));
            return;
            }
            this._contentLoaded = true;
        }
    }
}

