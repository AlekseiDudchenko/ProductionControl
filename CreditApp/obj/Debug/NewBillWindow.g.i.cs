﻿#pragma checksum "..\..\NewBillWindow.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "65E2CE85931AB509F08F69072B83A4E9"
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
    /// NewBillWindow
    /// </summary>
    public partial class NewBillWindow : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 26 "..\..\NewBillWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox NomberBillTextBox;
        
        #line default
        #line hidden
        
        
        #line 32 "..\..\NewBillWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox ProviderNameComboBox;
        
        #line default
        #line hidden
        
        
        #line 38 "..\..\NewBillWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox BillPriceTextBox;
        
        #line default
        #line hidden
        
        
        #line 43 "..\..\NewBillWindow.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button CreateNewBillButton;
        
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
            System.Uri resourceLocater = new System.Uri("/CreditApp;component/newbillwindow.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\NewBillWindow.xaml"
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
            this.NomberBillTextBox = ((System.Windows.Controls.TextBox)(target));
            
            #line 28 "..\..\NewBillWindow.xaml"
            this.NomberBillTextBox.TextChanged += new System.Windows.Controls.TextChangedEventHandler(this.NomberBillTextBox_TextChanged);
            
            #line default
            #line hidden
            return;
            case 2:
            this.ProviderNameComboBox = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 3:
            this.BillPriceTextBox = ((System.Windows.Controls.TextBox)(target));
            
            #line 41 "..\..\NewBillWindow.xaml"
            this.BillPriceTextBox.TextChanged += new System.Windows.Controls.TextChangedEventHandler(this.BillPriceTextBox_TextChanged);
            
            #line default
            #line hidden
            return;
            case 4:
            this.CreateNewBillButton = ((System.Windows.Controls.Button)(target));
            
            #line 44 "..\..\NewBillWindow.xaml"
            this.CreateNewBillButton.Click += new System.Windows.RoutedEventHandler(this.CreateNewBill);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

