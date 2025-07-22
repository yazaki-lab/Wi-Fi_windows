using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Windows.Forms;
using NativeWifi;

namespace WiFiAnalyzer
{
    public partial class MainForm : Form
    {
        private WlanClient wlanClient;
        private WlanClient.WlanInterface wlanInterface;
        
        // UI コンポーネント
        private Label currentSSIDLabel;
        private Label currentBSSIDLabel;
        private Button scanButton;
        private DataGridView networksDataGridView;
        
        public MainForm()
        {
            InitializeComponent();
            InitializeWiFi();
        }
        
        private void InitializeComponent()
        {
            this.Text = "Wi-Fi Analyzer";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            
            // 現在のSSIDラベル
            currentSSIDLabel = new Label
            {
                Text = "現在のSSID: 取得中...",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(400, 25),
                Font = new System.Drawing.Font("メイリオ", 10F)
            };
            this.Controls.Add(currentSSIDLabel);
            
            // 現在のBSSIDラベル
            currentBSSIDLabel = new Label
            {
                Text = "現在のBSSID: 取得中...",
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(400, 25),
                Font = new System.Drawing.Font("メイリオ", 10F)
            };
            this.Controls.Add(currentBSSIDLabel);
            
            // スキャンボタン
            scanButton = new Button
            {
                Text = "Wi-Fiスキャン",
                Location = new System.Drawing.Point(20, 85),
                Size = new System.Drawing.Size(120, 35),
                Font = new System.Drawing.Font("メイリオ", 10F)
            };
            scanButton.Click += ScanButton_Click;
            this.Controls.Add(scanButton);
            
            // データグリッドビュー
            networksDataGridView = new DataGridView
            {
                Location = new System.Drawing.Point(20, 130),
                Size = new System.Drawing.Size(740, 420),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new System.Drawing.Font("メイリオ", 9F)
            };
            
            // カラムの設定
            networksDataGridView.Columns.Add("SSID", "SSID");
            networksDataGridView.Columns.Add("BSSID", "BSSID");
            networksDataGridView.Columns.Add("Signal", "信号強度 (%)");
            networksDataGridView.Columns.Add("RSSI", "RSSI (dBm)");
            networksDataGridView.Columns.Add("Channel", "チャンネル");
            networksDataGridView.Columns.Add("Security", "セキュリティ");
            
            // カラム幅の設定
            networksDataGridView.Columns["SSID"].Width = 200;
            networksDataGridView.Columns["BSSID"].Width = 150;
            networksDataGridView.Columns["Signal"].Width = 100;
            networksDataGridView.Columns["RSSI"].Width = 100;
            networksDataGridView.Columns["Channel"].Width = 80;
            networksDataGridView.Columns["Security"].Width = 110;
            
            this.Controls.Add(networksDataGridView);
        }
        
        private void InitializeWiFi()
        {
            try
            {
                wlanClient = new WlanClient();
                if (wlanClient.Interfaces.Length > 0)
                {
                    wlanInterface = wlanClient.Interfaces[0];
                    UpdateCurrentConnectionInfo();
                }
                else
                {
                    currentSSIDLabel.Text = "現在のSSID: Wi-Fiアダプタが見つかりません";
                    currentBSSIDLabel.Text = "現在のBSSID: Wi-Fiアダプタが見つかりません";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Wi-Fi初期化エラー: {ex.Message}", "エラー", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void UpdateCurrentConnectionInfo()
        {
            try
            {
                if (wlanInterface != null && wlanInterface.InterfaceState == Wlan.WlanInterfaceState.Connected)
                {
                    var connectedProfile = wlanInterface.CurrentConnection;
                    currentSSIDLabel.Text = $"現在のSSID: {connectedProfile.profileName}";
                    
                    // BSSIDの取得は複雑なため、接続中であることのみ表示
                    currentBSSIDLabel.Text = "現在のBSSID: 接続中";
                }
                else
                {
                    currentSSIDLabel.Text = "現在のSSID: 未接続";
                    currentBSSIDLabel.Text = "現在のBSSID: 未接続";
                }
            }
            catch (Exception ex)
            {
                currentSSIDLabel.Text = $"現在のSSID: 取得エラー ({ex.Message})";
                currentBSSIDLabel.Text = "現在のBSSID: 取得エラー";
            }
        }
        
        private async void ScanButton_Click(object sender, EventArgs e)
        {
            if (wlanInterface == null)
            {
                MessageBox.Show("Wi-Fiアダプタが利用できません。", "エラー", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            scanButton.Enabled = false;
            scanButton.Text = "スキャン中...";
            networksDataGridView.Rows.Clear();
            
            try
            {
                await Task.Run(() => PerformWiFiScan());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"スキャンエラー: {ex.Message}", "エラー", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                scanButton.Enabled = true;
                scanButton.Text = "Wi-Fiスキャン";
                UpdateCurrentConnectionInfo();
            }
        }
        
        private void PerformWiFiScan()
        {
            try
            {
                wlanInterface.Scan();
                System.Threading.Thread.Sleep(3000); // スキャン完了を待つ
                
                var availableNetworks = wlanInterface.GetAvailableNetworkList(0);
                var networkList = new List<NetworkInfo>();
                
                foreach (var network in availableNetworks)
                {
                    var networkInfo = new NetworkInfo
                    {
                        SSID = GetStringForSSID(network.dot11Ssid),
                        BSSID = "N/A", // NativeWifiライブラリでは直接取得困難
                        SignalQuality = (int)network.wlanSignalQuality,
                        RSSI = ConvertSignalQualityToRSSI(network.wlanSignalQuality),
                        Channel = "N/A", // 追加の処理が必要
                        Security = network.dot11DefaultAuthAlgorithm.ToString()
                    };
                    networkList.Add(networkInfo);
                }
                
                // 信号強度でソート
                networkList = networkList.OrderByDescending(n => n.SignalQuality).ToList();
                
                // UIスレッドでDataGridViewを更新
                this.Invoke(new Action(() =>
                {
                    foreach (var network in networkList)
                    {
                        networksDataGridView.Rows.Add(
                            string.IsNullOrEmpty(network.SSID) ? "(Hidden)" : network.SSID,
                            network.BSSID,
                            $"{network.SignalQuality}%",
                            $"{network.RSSI} dBm",
                            network.Channel,
                            network.Security
                        );
                    }
                }));
            }
            catch (Exception ex)
            {
                this.Invoke(new Action(() =>
                {
                    MessageBox.Show($"スキャン処理エラー: {ex.Message}", "エラー", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
            }
        }
        
        private string GetStringForSSID(Wlan.Dot11Ssid ssid)
        {
            return System.Text.Encoding.UTF8.GetString(ssid.SSID, 0, (int)ssid.SSIDLength);
        }
        
        private int ConvertSignalQualityToRSSI(uint signalQuality)
        {
            // 近似値でSignal QualityからRSSIに変換
            // 実際の変換はより複雑
            return (int)(signalQuality / 2) - 100;
        }
        
        private class NetworkInfo
        {
            public string SSID { get; set; }
            public string BSSID { get; set; }
            public int SignalQuality { get; set; }
            public int RSSI { get; set; }
            public string Channel { get; set; }
            public string Security { get; set; }
        }
    }
    
    // Program.cs
    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}