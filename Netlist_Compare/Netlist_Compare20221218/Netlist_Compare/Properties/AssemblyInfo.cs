using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// 組件的一般資訊是由下列的屬性集控制。
// 變更這些屬性的值即可修改組件的相關
// 資訊。
[assembly: AssemblyTitle("Netlist_Compare")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("Netlist_Compare")]
[assembly: AssemblyCopyright("Copyright ©  2022")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// 將 ComVisible 設為 false 可對 COM 元件隱藏
// 組件中的類型。若必須從 COM 存取此組件中的類型，
// 的類型，請在該類型上將 ComVisible 屬性設定為 true。
[assembly: ComVisible(false)]

// 下列 GUID 為專案公開 (Expose) 至 COM 時所要使用的 typelib ID
[assembly: Guid("2c8c03cc-feb4-4ab9-bfda-c3a9db735f5d")]

// 組件的版本資訊由下列四個值所組成: 
//
//      主要版本
//      次要版本
//      組建編號
//      修訂編號
//
// 您可以指定所有的值，也可以使用 '*' 將組建和修訂編號
// 設為預設，如下所示:
 //[assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("2022.09.04.1")]
[assembly: AssemblyFileVersion("2022.09.04.1")] //兩者要同時進版
/*
                "\n 版本: 2022.08.07.1 初版" +
                "\n 版本: 2022.08.20.1 版本控制" +
                "\n 版本: 2022.08.20.2 比過的不重複比較" +
                "\n 版本: 2022.08.21.1" +
                "\n        1.) New net, package 只保留有差異的部分" +
                "\n        2.) Compare only run 1 times" +
                "\n        3.) 修正Net排序不同，判斷為【差異】結果問題" +  
                "\n 版本: 2022.09.04.1" +
                "\n        1.) 排除Net排序不同判斷差異" +
                "\n        2.) 新增Sky_Mode" +
                "\n        3.) 新增只有[Net Name]差異 => [橘色]" +
                  
                "\n\n 待更新版本: 2022.09.??.1 " +
                "\n1.) 下次修改程式運行速度" +
                "\n2.) 進階比較，Net Name不同，netlist只有些微不同" +
                "\n2.) 軟體自動更新","版本說明");
*/
