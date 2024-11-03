using ExcelFileInspector.Utilities;
using Prism.Commands;
using Prism.Mvvm;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Reflection;

namespace ExcelFileInspector.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        //--------------------------------------------------
        // バインディングデータ
        //--------------------------------------------------
        /// <summary>
        /// タイトル
        /// </summary>
        private string _title = Resources.Strings.ApplicationName;
        public string Title
        {
            get { return _title; }
            set { SetProperty(ref _title, value); }
        }

        /// <summary>
        /// コピーライト
        /// </summary>
        private string _copyright = string.Empty;
        public string Copyright
        {
            get { return _copyright; }
            set { SetProperty(ref _copyright, value); }
        }

        /// <summary>
        /// 検査設定プリセット(リスト)
        /// </summary>
        private ObservableCollection<string> _inspectionSettingPresetList = [];
        public ObservableCollection<string> InspectionSettingPresetList
        {
            get { return _inspectionSettingPresetList; }
            set { SetProperty(ref _inspectionSettingPresetList, value); }
        }

        /// <summary>
        /// プリセット(選択状態)
        /// </summary>
        private string _inspectionSettingPreset = string.Empty;
        public string InspectionSettingPreset
        {
            get { return _inspectionSettingPreset; }
            set { SetProperty(ref _inspectionSettingPreset, value); }
        }

        /// <summary>
        /// 対象ファイルキーワード
        /// </summary>
        private string _targetFileKeyword = string.Empty;
        public string TargetFileKeyword
        {
            get { return _targetFileKeyword; }
            set { SetProperty(ref _targetFileKeyword, value); }
        }

        /// <summary>
        /// 検査方法(リスト)
        /// </summary>
        private ObservableCollection<SettingManager.InspectionMethod> _inspectionMethods = [];
        public ObservableCollection<SettingManager.InspectionMethod> InspectionMethods
        {
            get { return _inspectionMethods; }
            set { SetProperty(ref _inspectionMethods, value); }
        }

        /// <summary>
        /// 検査ファイル
        /// </summary>
        private ObservableCollection<string> _inspectionFiles = [];
        public ObservableCollection<string> InspectionFiles
        {
            get { return _inspectionFiles; }
            set { SetProperty(ref _inspectionFiles, value); }
        }

        /// <summary>
        /// 検査結果
        /// </summary>
        private ObservableCollection<Inspector.InspectionResult> _inspectionResult = [];
        public ObservableCollection<Inspector.InspectionResult> InspectionResult
        {
            get { return _inspectionResult; }
            set { SetProperty(ref _inspectionResult, value); }
        }

        /// <summary>
        /// 進捗メッセージ
        /// </summary>
        private string _progressMessage = string.Empty;
        public string ProgressMessage
        {
            get { return _progressMessage; }
            set { SetProperty(ref _progressMessage, value); }
        }

        /// <summary>
        /// 操作可能フラグ
        /// </summary>
        private bool _isOperationEnable = true;
        public bool IsOperationEnable
        {
            get { return _isOperationEnable; }
            set { SetProperty(ref _isOperationEnable, value); }
        }

        //--------------------------------------------------
        // バインディングコマンド
        //--------------------------------------------------
        /// <summary>
        /// 検査実施
        /// </summary>
        private DelegateCommand _commandInspection;
        public DelegateCommand CommandInspection => _commandInspection ??= new DelegateCommand(ExecuteCommandInspection);

        /// <summary>
        /// URLを開く
        /// </summary>
        private DelegateCommand<string> _commandOpenUrl;
        public DelegateCommand<string> CommandOpenUrl => _commandOpenUrl ??= new DelegateCommand<string>(ExecuteCommandOpenUrl);

        //--------------------------------------------------
        // メソッド
        //--------------------------------------------------
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainWindowViewModel()
        {
            Assembly assm = Assembly.GetExecutingAssembly();
            string version = assm.GetCustomAttribute<AssemblyInformationalVersionAttribute>().InformationalVersion;

            // バージョン情報を取得してタイトルに反映する
            Title = $"{Resources.Strings.ApplicationName} Ver.{version}";

            // コピーライト情報を取得して設定する
            Copyright = assm.GetCustomAttribute<AssemblyCopyrightAttribute>().Copyright;
        }

        /// <summary>
        /// 検査実施コマンド実行処理
        /// </summary>
        private void ExecuteCommandInspection()
        {

        }

        /// <summary>
        /// URLを開くコマンド実行処理
        /// </summary>
        private void ExecuteCommandOpenUrl(string url)
        {
            ProcessStartInfo psi = new() { FileName = url, UseShellExecute = true };
            Process.Start(psi);
        }
    }
}
