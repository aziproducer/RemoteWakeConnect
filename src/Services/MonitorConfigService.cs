using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using RemoteWakeConnect.Models;

namespace RemoteWakeConnect.Services
{
    public class MonitorConfigService
    {
        /// <summary>
        /// 現在のモニター構成のハッシュ値を生成
        /// </summary>
        public string GenerateMonitorConfigHash(List<MonitorInfo> monitors)
        {
            if (monitors == null || monitors.Count == 0)
                return string.Empty;

            // モニター情報を文字列化（位置、サイズ、プライマリ情報を含む）
            var configString = string.Join("|", monitors
                .OrderBy(m => m.X)
                .ThenBy(m => m.Y)
                .Select(m => $"{m.X},{m.Y},{m.Width},{m.Height},{m.IsPrimary}"));

            // ハッシュ値を生成
            using (var sha256 = SHA256.Create())
            {
                var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(configString));
                return Convert.ToBase64String(bytes);
            }
        }

        /// <summary>
        /// モニター構成が変更されているかチェック
        /// </summary>
        public bool HasMonitorConfigChanged(string savedHash, List<MonitorInfo> currentMonitors)
        {
            if (string.IsNullOrEmpty(savedHash))
                return false; // 初回は変更なしとする

            var currentHash = GenerateMonitorConfigHash(currentMonitors);
            return savedHash != currentHash;
        }

        /// <summary>
        /// 保存されたモニター選択を復元
        /// </summary>
        public void RestoreMonitorSelection(List<MonitorInfo> monitors, List<int> savedIndices)
        {
            // すべてのモニターの選択を解除
            foreach (var monitor in monitors)
            {
                monitor.IsSelected = false;
            }

            // 保存されたインデックスのモニターを選択
            foreach (var index in savedIndices)
            {
                if (index >= 0 && index < monitors.Count)
                {
                    monitors[index].IsSelected = true;
                }
            }
        }

        /// <summary>
        /// 現在の選択されたモニターのインデックスリストを取得
        /// </summary>
        public List<int> GetSelectedMonitorIndices(List<MonitorInfo> monitors)
        {
            return monitors
                .Where(m => m.IsSelected)
                .Select(m => m.Index)
                .ToList();
        }

        /// <summary>
        /// モニター構成の変更内容を説明する文字列を生成
        /// </summary>
        public string GetMonitorConfigChangeDescription(int savedCount, int currentCount)
        {
            if (savedCount == currentCount)
            {
                return "モニターの配置または解像度が変更されています。";
            }
            else if (savedCount < currentCount)
            {
                return $"モニターが追加されています。（{savedCount}台 → {currentCount}台）";
            }
            else
            {
                return $"モニターが減少しています。（{savedCount}台 → {currentCount}台）";
            }
        }
    }
}