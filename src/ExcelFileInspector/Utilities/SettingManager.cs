﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace ExcelFileInspector.Utilities
{
    /// <summary>
    /// 設定管理クラス
    /// </summary>
    public class SettingManager
    {
        /// <summary>
        /// 設定クラス
        /// </summary>
        public class Setting
        {
            /// <summary>
            /// プリセット名
            /// </summary>
            public string PresetName { get; set; } = string.Empty;

            /// <summary>
            /// 対象ファイルキーワード
            /// </summary>
            public string TargetFileKeyword { get; set; } = string.Empty;

            /// <summary>
            /// 検査方法
            /// </summary>
            public List<InspectionMethod> InspectionMethods { get; set; } = [];
        }

        /// <summary>
        /// 検査方法クラス
        /// </summary>
        public class InspectionMethod
        {
            /// <summary>
            /// シート名
            /// </summary>
            public string SheetName { get; set; } = string.Empty;

            /// <summary>
            /// セル
            /// </summary>
            public string Cell { get; set; } = string.Empty;

            /// <summary>
            /// 条件
            /// </summary>
            public string Condition { get; set; } = string.Empty;

            /// <summary>
            /// 値
            /// </summary>
            public string Value { get; set; } = string.Empty;
        }

        //--------------------------------------------------
        // 定数(コンフィギュレーション)
        //--------------------------------------------------
        /// <summary>
        /// 設定ファイル名
        /// </summary>
        private static readonly string _fileName = "InspectionSettings.json";

        /// <summary>
        /// デシリアライズ設定(コメント無視)
        /// </summary>
        private static readonly JsonSerializerOptions _deserializeOptions = new() { ReadCommentHandling = JsonCommentHandling.Skip };

        /// <summary>
        /// シリアライズ設定(インデントあり/日本語ありのためエンコーダ設定/高速化のためUTF-8 バイトの配列にシリアル化)
        /// </summary>
        private static readonly JsonSerializerOptions _serializeOptions = new() { WriteIndented = true, Encoder = JavaScriptEncoder.Create(UnicodeRanges.All) };

        //--------------------------------------------------
        // 内部変数
        //--------------------------------------------------
        /// <summary>
        /// 設定
        /// </summary>
        private static List<Setting> _settings = [];

        //--------------------------------------------------
        // メソッド
        //--------------------------------------------------
        /// <summary>
        /// 設定ファイル読み込み処理
        /// </summary>
        public static void ReadSettings()
        {
            string settingFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _fileName);

            // 設定ファイルがない場合は新規作成する
            if (!File.Exists(settingFilePath))
            {
                WriteSettings();
            }

            // 設定ファイルの読み込み
            string jsonString = File.ReadAllText(_fileName);

            // デシリアライズ
            _settings = JsonSerializer.Deserialize<List<Setting>>(jsonString, _deserializeOptions);
        }

        /// <summary>
        /// 設定ファイル書き込み処理
        /// </summary>
        private static void WriteSettings()
        {
            string settingFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _fileName);

            // シリアライズ
            byte[] jsonUtf8Bytes = JsonSerializer.SerializeToUtf8Bytes(_settings, _serializeOptions);

            // ファイル出力
            using FileStream fs = new(settingFilePath, FileMode.Create);
            fs.Write(jsonUtf8Bytes);
            fs.Close();
        }

        /// <summary>
        /// プリセットリスト取得処理
        /// </summary>
        /// <returns>プリセットリスト</returns>
        public static List<string> GetPresetList()
        {
            List<string> presetList = [];

            // 読み込み済み設定からプリセット名を取得してプリセットリストを生成する
            foreach (Setting item in _settings)
            {
                presetList.Add(item.PresetName);
            }

            return presetList;
        }

        /// <summary>
        /// プリセット取得処理
        /// </summary>
        /// <param name="presetName">プリセット名</param>
        /// <returns>プリセット</returns>
        public static Setting GetPreset(string presetName)
        {
            Setting setting = new();

            // 読み込み済み設定からプリセット名が一致するものを返す
            foreach (Setting item in _settings)
            {
                if(item.PresetName == presetName)
                {
                    setting = item;
                }
            }

            return setting;
        }
    }
}
