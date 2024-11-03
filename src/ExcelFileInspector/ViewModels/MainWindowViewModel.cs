using ExcelFileInspector.Utilities;
using Prism.Commands;
using Prism.Mvvm;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;

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
        /// 検査設定プリセット(選択状態)
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
        private ObservableCollection<Inspector.InspectionResult> _inspectionResultList = [];
        public ObservableCollection<Inspector.InspectionResult> InspectionResultList
        {
            get { return _inspectionResultList; }
            set { SetProperty(ref _inspectionResultList, value); }
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
        /// プログレスバー最大値
        /// </summary>
        private int _progressMaximum = 1;
        public int ProgressMaximum
        {
            get { return _progressMaximum; }
            set { SetProperty(ref _progressMaximum, value); }
        }

        /// <summary>
        /// プログレスバー現在値
        /// </summary>
        private int _progressValue = 0;
        public int ProgressValue
        {
            get { return _progressValue; }
            set { SetProperty(ref _progressValue, value); }
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
        /// プリセット選択変更
        /// </summary>
        private DelegateCommand _commandPresetChange;
        public DelegateCommand CommandPresetChange => _commandPresetChange ??= new DelegateCommand(ExecuteCommandPresetChange);

        /// <summary>
        /// 検査ファイルクリア
        /// </summary>
        private DelegateCommand _commandClearInspectionFile;
        public DelegateCommand CommandClearInspectionFile => _commandClearInspectionFile ??= new DelegateCommand(ExecuteCommandClearInspectionFile);

        /// <summary>
        /// 検査ファイルドラッグ
        /// </summary>
        private DelegateCommand<DragEventArgs> _commandInspectionPreviewDragOver;
        public DelegateCommand<DragEventArgs> CommandInspectionPreviewDragOver => _commandInspectionPreviewDragOver ??= new DelegateCommand<DragEventArgs>(ExecuteCommandInspectionPreviewDragOver);

        /// <summary>
        /// 検査ファイルドロップ
        /// </summary>
        private DelegateCommand<DragEventArgs> _commandInspectionFileDrop;
        public DelegateCommand<DragEventArgs> CommandInspectionFileDrop => _commandInspectionFileDrop ??= new DelegateCommand<DragEventArgs>(ExecuteCommandInspectionFileDrop);

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

            if (!(bool)System.ComponentModel.DesignerProperties.IsInDesignModeProperty.GetMetadata(typeof(System.Windows.DependencyObject)).DefaultValue)
            {
                // 設定ファイルの読み込み
                // (XAMLデザイナーのエラー対策でデザインモードではない場合のみ)
                SettingManager.ReadSettings();
            }
            
            // 検査設定プリセット(リスト)の作成
            InspectionSettingPresetList = new(SettingManager.GetPresetList());

            // 検査設定の反映(プリセットリスト先頭)
            if (InspectionSettingPresetList.Count != 0)
            {
                LoadInspectionSettings(InspectionSettingPresetList[0]);
            }

            // 検査実施できるか確認
            CheckExecuteInspection();
        }

        /// <summary>
        /// 検査実施コマンド実行処理
        /// </summary>
        private async void ExecuteCommandInspection()
        {
            List<Inspector.InspectionResult> inspectionResult = [];
            List<SettingManager.InspectionMethod> inspectionMethods = new(InspectionMethods);

            // 検査開始
            InspectionResultList.Clear();
            ProgressMaximum = InspectionFiles.Count;
            ProgressValue = 0;
            IsOperationEnable = false;
            ProgressMessage = string.Format(Resources.Strings.MessageStatusNowInspection, ProgressValue, ProgressMaximum);

            // 検査実施
            await Task.Run(() =>
            {
                foreach (string file in InspectionFiles)
                {
                    List<Inspector.InspectionResult> inspectionResults = Inspector.InspectionFile(file, inspectionMethods);
                    foreach (Inspector.InspectionResult result in inspectionResults)
                    {
                        inspectionResult.Add(result);
                    }
                    ProgressValue++;
                    ProgressMessage = string.Format(Resources.Strings.MessageStatusNowInspection, ProgressValue, ProgressMaximum);
                }
            });

            // 検査完了
            InspectionResultList = new(inspectionResult);
            IsOperationEnable = true;
            ProgressMessage = Resources.Strings.MessageStatusCompleteInspection;
        }

        /// <summary>
        /// プリセット選択変更コマンド実行処理
        /// </summary>
        void ExecuteCommandPresetChange()
        {
            // 検査設定の反映
            LoadInspectionSettings(InspectionSettingPreset);
        }

        /// <summary>
        /// 検査ファイルクリアコマンド実行処理
        /// </summary>
        void ExecuteCommandClearInspectionFile()
        {
            // 検査ファイルリストをクリア
            InspectionFiles.Clear();

            // 検査実施できるか確認
            CheckExecuteInspection();
        }

        /// <summary>
        /// 検査ファイルドラッグコマンド実行処理
        /// </summary>
        /// <param name="e">イベントデータ</param>
        void ExecuteCommandInspectionPreviewDragOver(DragEventArgs e)
        {
            // ドラッグしてきたデータがファイルの場合､ドロップを許可する｡
            e.Effects = DragDropEffects.Copy;
            e.Handled = e.Data.GetDataPresent(DataFormats.FileDrop);
        }

        /// <summary>
        /// 検査ファイルドロップコマンド実行処理
        /// </summary>
        /// <param name="e">イベントデータ</param>
        private void ExecuteCommandInspectionFileDrop(DragEventArgs e)
        {
            // ドロップしてきたデータを解析する
            if (e.Data.GetData(DataFormats.FileDrop) is string[] dropitems)
            {
                foreach (string dropitem in dropitems)
                {
                    if (System.IO.Directory.Exists(dropitem) == true)
                    {
                        // フォルダの場合は配下の対象ファイルキーワードを含むサポートExcelファイルを検査ファイルのリストに追加
                        if (System.IO.Directory.GetFiles(@dropitem, "*", System.IO.SearchOption.AllDirectories) is string[] files)
                        {
                            foreach (string file in files)
                            {
                                if (file.Contains(TargetFileKeyword) && (System.IO.Path.GetExtension(file) == ".xlsx" || System.IO.Path.GetExtension(file) == ".xlsm"))
                                {
                                    InspectionFiles.Add(file);
                                }
                            }
                        }
                    }
                    else
                    {
                        // 対象ファイルキーワードを含むサポートExcelファイルの場合は検査ファイルのリストに追加
                        if (dropitem.Contains(TargetFileKeyword) && (System.IO.Path.GetExtension(dropitem) == ".xlsx" || System.IO.Path.GetExtension(dropitem) == ".xlsm"))
                        {
                            InspectionFiles.Add(dropitem);
                        }
                    }
                }
            }

            // 検査実施できるか確認
            CheckExecuteInspection();
        }

        /// <summary>
        /// URLを開くコマンド実行処理
        /// </summary>
        private void ExecuteCommandOpenUrl(string url)
        {
            ProcessStartInfo psi = new() { FileName = url, UseShellExecute = true };
            Process.Start(psi);
        }

        /// <summary>
        /// 検査設定反映処理
        /// </summary>
        /// <param name="presetName">プリセット名</param>
        private void LoadInspectionSettings(string presetName)
        {
            // プリセット取得
            SettingManager.Setting setting = SettingManager.GetPreset(presetName);

            // 検査設定プリセット(選択状態)に反映
            InspectionSettingPreset = presetName;

            // 対象ファイルキーワードに反映
            TargetFileKeyword = setting.TargetFileKeyword;

            // 検査方法(リスト)に反映
            InspectionMethods = new(setting.InspectionMethods);

            // 検査実施できるか確認
            CheckExecuteInspection();
        }

        /// <summary>
        /// 検査実施可否確認処理
        /// </summary>
        private void CheckExecuteInspection()
        {
            if (InspectionMethods.Count == 0)
            {
                // 検査設定の検査方法がない場合は検査実施不可
                IsOperationEnable = false;
                ProgressMessage = Resources.Strings.MessageStatusInspectionMethodEmpty;
            }
            else if(InspectionFiles.Count == 0)
            {
                // 検査ファイルがない場合は検査実施不可
                IsOperationEnable = false;
                ProgressMessage = Resources.Strings.MessageStatusInspectionFileEmpty;
            }
            else
            {
                // チェック通過時検査実施可能
                IsOperationEnable = true;
                ProgressMessage = Resources.Strings.MessageStatusAlreadyInspection;
            }
        }
    }
}
