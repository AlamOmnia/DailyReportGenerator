﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Demo_Excel_Export.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("Demo_Excel_Export.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to use purple;
        ///select c.partnername as &apos;SourceNetwork&apos;,CallCount AS &apos;CallCount&apos;,ActualDuration AS &apos;ActualDuration&apos;,BilledDuration AS &apos;BilledDuration&apos; from
        ///(
        ///select customerid,supplierid, DATE_FORMAT(date(starttime),&apos;%d/%m/%Y&apos;) `Date`,Count(*) as CallCount,sum(durationsec)/60 ActualDuration,
        ///sum(case when (truncate(durationsec-truncate(durationsec,0),1))&gt;=0
        ///then ceiling(durationsec)
        ///else floor(durationsec) end)/60 as RoundedDuration,sum(Duration1)/60 as BilledDuration
        ///from purple.cdrloaded
        ///where calldirection=1 [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string Dom_Monthly {
            get {
                return ResourceManager.GetString("Dom_Monthly", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Asia Alliance	x
        ///Apple Networks Ltd	x
        ///Bangla Tel ltd	x
        ///BanglaTrac	Btrac
        ///Bestec Telecom Ltd.	x
        ///BG Tel	x
        ///BIGL	x
        ///BTCL	x
        ///Cel Telecom Ltd	x
        ///DBL Telecom Ltd	x
        ///Digicon Telecommunciation	DIGICON
        ///First Communication Ltd	x
        ///Global Voice Telecom	GVTEL
        ///Hamid Sourcing Ltd.	x
        ///HRC Technologies	x
        ///Kay Telecommu nications Ltd	x
        ///Mir telecom	MIR
        ///Mos5Tel Ltd	x
        ///NovoTel	novotel
        ///Platinum Communication ltd	x
        ///Ranks Tel	x
        ///Ratul Telecom	x
        ///Roots Communications	Roots
        ///Sigma Eneineers Ltd.	x
        ///SM Communication	x
        ///Telex [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string IgwExcelDisplayToSourceNetworkMapping {
            get {
                return ResourceManager.GetString("IgwExcelDisplayToSourceNetworkMapping", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to use purple;
        ///select c.partnername as &apos;SourceNetwork&apos;,CallCount AS &apos;CallCount&apos;,ActualDuration AS &apos;ActualDuration&apos;,BilledDuration AS &apos;BilledDuration&apos; from
        ///(
        ///select customerid,SUM(SequenceNumber)CallCount ,sum(durationsec)/60 ActualDuration,
        ///sum(Duration1)/60 as BilledDuration
        ///from purple.cdrsummary
        ///where calldirection=3
        ///and starttime&gt;=@startTime
        ///and starttime&lt;@endTime
        ///group by CustomerID
        ///) x
        ///left join
        ///partner c
        ///on x.customerid=c.idpartner;.
        /// </summary>
        internal static string Int_Incom_IOS_Wise {
            get {
                return ResourceManager.GetString("Int_Incom_IOS_Wise", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to use purple;
        ///select c.partnername as &apos;SourceNetwork&apos;,CallCount AS &apos;CallCount&apos;,ActualDuration AS &apos;ActualDuration&apos;,BilledDuration AS &apos;BilledDuration&apos; from
        ///(
        ///select supplierid,SUM(SequenceNumber) CallCount,sum(durationsec)/60 ActualDuration,
        ///sum(roundedduration)/60 as BilledDuration
        ///from purple.cdrsummary
        ///where calldirection=2
        ///and starttime&gt;=@startTime
        ///and starttime&lt;@endTime
        ///group by supplierid
        ///) x
        ///left join
        ///partner c
        ///on x.supplierid=c.idpartner
        ///
        ///.
        /// </summary>
        internal static string Int_Outgoing_IOS_Wise {
            get {
                return ResourceManager.GetString("Int_Outgoing_IOS_Wise", resourceCulture);
            }
        }
    }
}