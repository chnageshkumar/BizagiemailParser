﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace takeda.bizagi.connector.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "15.3.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("http://des05346.nycomed.local/CBT/webservices/entitymanagersoa.asmx")]
        public string BizAgiConnectorLibrary_EntityManagerSOA_EntityManagerSOA {
            get {
                return ((string)(this["BizAgiConnectorLibrary_EntityManagerSOA_EntityManagerSOA"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("http://des05346.nycomed.local/Play/webservices/cache.asmx")]
        public string BizAgiConnectorLibrary_BizAgiCacheWebservice_Cache {
            get {
                return ((string)(this["BizAgiConnectorLibrary_BizAgiCacheWebservice_Cache"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("http://des05346.nycomed.local/FAP/webservices/workflowenginesoa.asmx")]
        public string BizAgiConnectorLibrary_WorkflowEngineSOA_WorkflowEngineSOA {
            get {
                return ((string)(this["BizAgiConnectorLibrary_WorkflowEngineSOA_WorkflowEngineSOA"]));
            }
        }
    }
}
