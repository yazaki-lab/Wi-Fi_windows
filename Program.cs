using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using ManagedNativeWifi;

namespace WiFiAnalyzer
{
    public partial class MainForm : Form
    {
        // UI コンポーネント
        private Panel initialPanel;
        private Panel detailPanel;
        
        // 初期画面のコンポーネント
        private Label currentSSIDLabel;
        private Label currentRSSILabel;
        private Button getInfoButton;
        
        // 詳細画面のコンポーネント
        private Label detailSSIDLabel;
        private Label detailBSSIDLabel;
        private Label detailRSSILabel;
        private Label detailChannelLabel;
        private Label detailSecurityLabel;
        private Label strengthEvaluationLabel;
        private Label testResultLabel;
        private Button backButton;
        private Button fullScanButton;
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
            
            CreateInitialPanel();
            CreateDetailPanel();
            
            // 初期画面を表示
            ShowInitialPanel();
        }
        
        private void CreateInitialPanel()
        {
            initialPanel = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(800, 600),
                BackColor = Color.White
            };
            
            // タイトル
            Label titleLabel = new Label
            {
                Text = "Wi-Fi アナライザー",
                Location = new Point(300, 100),
                Size = new Size(200, 40),
                Font = new Font("メイリオ", 16F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };
            initialPanel.Controls.Add(titleLabel);
            
            // 現在のSSIDラベル
            currentSSIDLabel = new Label
            {
                Text = "現在のSSID: 取得中...",
                Location = new Point(200, 200),
                Size = new Size(400, 30),
                Font = new Font("メイリオ", 12F),
                TextAlign = ContentAlignment.MiddleCenter
            };
            initialPanel.Controls.Add(currentSSIDLabel);
            
            // 現在のRSSIラベル
            currentRSSILabel = new Label
            {
                Text = "RSSI: 取得中...",
                Location = new Point(200, 240),
                Size = new Size(400, 30),
                Font = new Font("メイリオ", 12F),
                TextAlign = ContentAlignment.MiddleCenter
            };
            initialPanel.Controls.Add(currentRSSILabel);
            
            // Wi-Fi情報取得ボタン
            getInfoButton = new Button
            {
                Text = "Wi-Fiの情報を取得しますか？",
                Location = new Point(300, 320),
                Size = new Size(200, 50),
                Font = new Font("メイリオ", 11F),
                BackColor = Color.LightBlue,
                FlatStyle = FlatStyle.Flat
            };
            getInfoButton.Click += GetInfoButton_Click;
            initialPanel.Controls.Add(getInfoButton);
            
            this.Controls.Add(initialPanel);
        }
        
        private void CreateDetailPanel()
        {
            detailPanel = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(800, 600),
                BackColor = Color.White,
                Visible = false
            };
            
            // 戻るボタン
            backButton = new Button
            {
                Text = "← 戻る",
                Location = new Point(20, 20),
                Size = new Size(80, 30),
                Font = new Font("メイリオ", 9F)
            };
            backButton.Click += BackButton_Click;
            detailPanel.Controls.Add(backButton);
            
            // タイトル
            Label detailTitleLabel = new Label
            {
                Text = "Wi-Fi 詳細情報",
                Location = new Point(300, 20),
                Size = new Size(200, 30),
                Font = new Font("メイリオ", 14F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };
            detailPanel.Controls.Add(detailTitleLabel);
            
            // 詳細情報ラベル群
            detailSSIDLabel = new Label
            {
                Text = "SSID: ",
                Location = new Point(50, 80),
                Size = new Size(300, 25),
                Font = new Font("メイリオ", 10F)
            };
            detailPanel.Controls.Add(detailSSIDLabel);
            
            detailBSSIDLabel = new Label
            {
                Text = "BSSID: ",
                Location = new Point(50, 110),
                Size = new Size(300, 25),
                Font = new Font("メイリオ", 10F)
            };
            detailPanel.Controls.Add(detailBSSIDLabel);
            
            detailRSSILabel = new Label
            {
                Text = "RSSI: ",
                Location = new Point(50, 140),
                Size = new Size(300, 25),
                Font = new Font("メイリオ", 10F)
            };
            detailPanel.Controls.Add(detailRSSILabel);
            
            detailChannelLabel = new Label
            {
                Text = "チャンネル: ",
                Location = new Point(50, 170),
                Size = new Size(300, 25),
                Font = new Font("メイリオ", 10F)
            };
            detailPanel.Controls.Add(detailChannelLabel);
            
            detailSecurityLabel = new Label
            {
                Text = "セキュリティ: ",
                Location = new Point(50, 200),
                Size = new Size(300, 25),
                Font = new Font("メイリオ", 10F)
            };
            detailPanel.Controls.Add(detailSecurityLabel);
            
            // 強度評価ラベル
            strengthEvaluationLabel = new Label
            {
                Text = "信号強度: 評価中...",
                Location = new Point(400, 80),
                Size = new Size(300, 40),
                Font = new Font("メイリオ", 12F, FontStyle.Bold),
                ForeColor = Color.Blue
            };
            detailPanel.Controls.Add(strengthEvaluationLabel);
            
            // 検査結果ラベル
            testResultLabel = new Label
            {
                Text = "検査結果: 分析中...",
                Location = new Point(400, 130),
                Size = new Size(300, 80),
                Font = new Font("メイリオ", 10F),
                ForeColor = Color.DarkGreen
            };
            detailPanel.Controls.Add(testResultLabel);
            
            // 全スキャンボタン
            fullScanButton = new Button
            {
                Text = "全Wi-Fiスキャン",
                Location = new Point(50, 250),
                Size = new Size(120, 35),
                Font = new Font("メイリオ", 10F)
            };
            fullScanButton.Click += FullScanButton_Click;
            detailPanel.Controls.Add(fullScanButton);
            
            // データグリッドビュー（全スキャン結果用）
            networksDataGridView = new DataGridView
            {
                Location = new Point(50, 300),
                Size = new Size(700, 250),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("メイリオ", 9F),
                Visible = false
            };
            
            // カラムの設定
            networksDataGridView.Columns.Add("SSID", "SSID");
            networksDataGridView.Columns.Add("BSSID", "BSSID");
            networksDataGridView.Columns.Add("Signal", "信号強度 (%)");
            networksDataGridView.Columns.Add("RSSI", "RSSI (dBm)");
            networksDataGridView.Columns.Add("Channel", "チャンネル");
            networksDataGridView.Columns.Add("Security", "セキュリティ");
            
            // カラム幅の設定
            networksDataGridView.Columns["SSID"].Width = 150;
            networksDataGridView.Columns["BSSID"].Width = 130;
            networksDataGridView.Columns["Signal"].Width = 90;
            networksDataGridView.Columns["RSSI"].Width = 90;
            networksDataGridView.Columns["Channel"].Width = 80;
            networksDataGridView.Columns["Security"].Width = 100;
            
            detailPanel.Controls.Add(networksDataGridView);
            
            this.Controls.Add(detailPanel);
        }
        
        private void InitializeWiFi()
        {
            try
            {
                UpdateCurrentConnectionInfo();
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
                var connectionProfile = NativeWifi.GetConnectionProfiles().FirstOrDefault(x => x.IsConnected);
                
                if (connectionProfile != null)
                {
                    currentSSIDLabel.Text = $"現在のSSID: {connectionProfile.Name}";
                    
                    try
                    {
                        // スキャンして現在の接続のRSSIを取得
                        var availableNetworks = NativeWifi.ScanNetworks();
                        var currentNetwork = availableNetworks.FirstOrDefault(x => 
                            x.Ssid == connectionProfile.Name);
                        
                        if (currentNetwork != null)
                        {
                            currentRSSILabel.Text = $"RSSI: {currentNetwork.SignalStrength} dBm";
                        }
                        else
                        {
                            currentRSSILabel.Text = "RSSI: 取得中...";
                        }
                    }
                    catch
                    {
                        currentRSSILabel.Text = "RSSI: 取得中...";
                    }
                }
                else
                {
                    currentSSIDLabel.Text = "現在のSSID: 未接続";
                    currentRSSILabel.Text = "RSSI: 未接続";
                }
            }
            catch (Exception ex)
            {
                currentSSIDLabel.Text = $"現在のSSID: 取得エラー ({ex.Message})";
                currentRSSILabel.Text = "RSSI: 取得エラー";
            }
        }
        
        private async void GetInfoButton_Click(object sender, EventArgs e)
        {
            getInfoButton.Enabled = false;
            getInfoButton.Text = "情報取得中...";
            
            try
            {
                await Task.Run(() => GetDetailedWiFiInfo());
                ShowDetailPanel();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"情報取得エラー: {ex.Message}", "エラー", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                getInfoButton.Enabled = true;
                getInfoButton.Text = "Wi-Fiの情報を取得しますか？";
            }
        }
        
        private void GetDetailedWiFiInfo()
        {
            try
            {
                var connectionProfile = NativeWifi.GetConnectionProfiles().FirstOrDefault(x => x.IsConnected);
                
                if (connectionProfile != null)
                {
                    var availableNetworks = NativeWifi.ScanNetworks();
                    var currentNetwork = availableNetworks.FirstOrDefault(x => 
                        x.Ssid == connectionProfile.Name);
                    
                    if (currentNetwork != null)
                    {
                        var networkInfo = new NetworkInfo
                        {
                            SSID = currentNetwork.Ssid,
                            BSSID = currentNetwork.BssId,
                            SignalQuality = CalculateSignalQuality(currentNetwork.SignalStrength),
                            RSSI = currentNetwork.SignalStrength,
                            Channel = GetChannelFromFrequency(currentNetwork.Frequency),
                            Security = GetSecurityType(currentNetwork)
                        };
                        
                        this.Invoke(new Action(() => UpdateDetailPanel(networkInfo)));
                    }
                    else
                    {
                        // 現在接続中のネットワークが見つからない場合のデフォルト情報
                        var networkInfo = new NetworkInfo
                        {
                            SSID = connectionProfile.Name,
                            BSSID = "取得できませんでした",
                            SignalQuality = 0,
                            RSSI = -100,
                            Channel = "不明",
                            Security = "不明"
                        };
                        
                        this.Invoke(new Action(() => UpdateDetailPanel(networkInfo)));
                    }
                }
                else
                {
                    this.Invoke(new Action(() =>
                    {
                        MessageBox.Show("現在Wi-Fiに接続されていません。", "情報", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                }
            }
            catch (Exception ex)
            {
                this.Invoke(new Action(() =>
                {
                    MessageBox.Show($"詳細情報取得エラー: {ex.Message}", "エラー", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
            }
        }
        
        private void UpdateDetailPanel(NetworkInfo networkInfo)
        {
            detailSSIDLabel.Text = $"SSID: {networkInfo.SSID}";
            detailBSSIDLabel.Text = $"BSSID: {networkInfo.BSSID}";
            detailRSSILabel.Text = $"RSSI: {networkInfo.RSSI} dBm ({networkInfo.SignalQuality}%)";
            detailChannelLabel.Text = $"チャンネル: {networkInfo.Channel}";
            detailSecurityLabel.Text = $"セキュリティ: {networkInfo.Security}";
            
            // 信号強度の評価
            string strength;
            Color strengthColor;
            
            if (networkInfo.RSSI >= -50)
            {
                strength = "強い (優秀)";
                strengthColor = Color.Green;
            }
            else if (networkInfo.RSSI >= -70)
            {
                strength = "中ぐらい (良好)";
                strengthColor = Color.Orange;
            }
            else
            {
                strength = "弱い (改善が必要)";
                strengthColor = Color.Red;
            }
            
            strengthEvaluationLabel.Text = $"信号強度: {strength}";
            strengthEvaluationLabel.ForeColor = strengthColor;
            
            // 検査結果の生成
            string testResult = GenerateTestResult(networkInfo);
            testResultLabel.Text = $"検査結果:\n{testResult}";
        }
        
        private string GenerateTestResult(NetworkInfo networkInfo)
        {
            var results = new List<string>();
            
            // RSSI評価
            if (networkInfo.RSSI >= -50)
                results.Add("✓ 信号強度は良好です");
            else if (networkInfo.RSSI >= -70)
                results.Add("△ 信号強度は普通です");
            else
                results.Add("✗ 信号強度が弱く、改善が必要です");
            
            // チャンネル評価（簡単な例）
            if (networkInfo.Channel == "1" || networkInfo.Channel == "6" || networkInfo.Channel == "11")
                results.Add("✓ 2.4GHz推奨チャンネルです");
            else if (int.TryParse(networkInfo.Channel, out int ch) && ch > 14)
                results.Add("✓ 5GHz帯を使用しています");
            else
                results.Add("△ チャンネル干渉の可能性があります");
            
            return string.Join("\n", results);
        }
        
        private async void FullScanButton_Click(object sender, EventArgs e)
        {
            fullScanButton.Enabled = false;
            fullScanButton.Text = "スキャン中...";
            networksDataGridView.Rows.Clear();
            networksDataGridView.Visible = true;
            
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
                fullScanButton.Enabled = true;
                fullScanButton.Text = "全Wi-Fiスキャン";
            }
        }
        
        private void PerformWiFiScan()
        {
            try
            {
                var availableNetworks = NativeWifi.ScanNetworks();
                var networkList = new List<NetworkInfo>();
                
                foreach (var network in availableNetworks)
                {
                    var networkInfo = new NetworkInfo
                    {
                        SSID = string.IsNullOrEmpty(network.Ssid) ? "(Hidden)" : network.Ssid,
                        BSSID = network.BssId,
                        SignalQuality = CalculateSignalQuality(network.SignalStrength),
                        RSSI = network.SignalStrength,
                        Channel = GetChannelFromFrequency(network.Frequency),
                        Security = GetSecurityType(network)
                    };
                    networkList.Add(networkInfo);
                }
                
                var uniqueNetworks = networkList
                    .GroupBy(n => n.SSID)
                    .Select(g => g.OrderByDescending(n => n.SignalQuality).First())
                    .OrderByDescending(n => n.SignalQuality)
                    .ToList();
                
                this.Invoke(new Action(() =>
                {
                    foreach (var network in uniqueNetworks)
                    {
                        networksDataGridView.Rows.Add(
                            network.SSID,
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
        
        private void BackButton_Click(object sender, EventArgs e)
        {
            ShowInitialPanel();
        }
        
        private void ShowInitialPanel()
        {
            initialPanel.Visible = true;
            detailPanel.Visible = false;
            getInfoButton.Enabled = true;
            getInfoButton.Text = "Wi-Fiの情報を取得しますか？";
        }
        
        private void ShowDetailPanel()
        {
            initialPanel.Visible = false;
            detailPanel.Visible = true;
            networksDataGridView.Visible = false;
        }
        
        private int CalculateSignalQuality(int rssi)
        {
            // RSSIから信号品質（%）を計算
            if (rssi >= -50) return 100;
            if (rssi <= -100) return 0;
            return 2 * (rssi + 100);
        }
        
        private string GetChannelFromFrequency(int frequency)
        {            
            // 2.4GHz帯
            if (frequency >= 2412 && frequency <= 2484)
            {
                if (frequency == 2484)
                    return "14";
                return ((frequency - 2412) / 5 + 1).ToString();
            }
            // 5GHz帯
            else if (frequency >= 5170 && frequency <= 5825)
            {
                return ((frequency - 5000) / 5).ToString();
            }
            
            return "N/A";
        }
        
        private string GetSecurityType(AvailableNetworkPack network)
        {
            // managedNativeWiFiでのセキュリティタイプの取得
            // ここは簡単な実装にします。より詳細な情報が必要な場合は
            // ライブラリのドキュメントを参照してください
            if (network.SignalStrength < -100)
                return "不明";
            
            // デフォルトでWPA2と仮定（実際の実装では詳細な判定が必要）
            return "WPA2";
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