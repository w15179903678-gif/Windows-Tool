using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Newtonsoft.Json;

namespace WindowTool
{
    /// <summary>
    /// 控制台主界面：集成连接、智能录制、循环回放及持久化功能
    /// </summary>
    public partial class MainWindow : Window
    {
        #region 私有字段

        private IntPtr _targetHwnd = IntPtr.Zero;
        private IntPtr _inputHwnd = IntPtr.Zero;
        private readonly AutomationTask _task = new();
        private readonly string _configDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Configs");

        // 录制相关状态
        private bool _isRecordingAttribute = false;
        private IntPtr _mouseHookHwnd = IntPtr.Zero;
        private DateTime _lastActionTimestamp;
        private Win32Helper.LowLevelMouseProc? _mouseProcDelegate;
        private Win32Helper.POINT _lastDownPosition;
        private DateTime _lastDownTimestamp;
        private bool _isLeftButtonActive = false;

        // 运行相关状态
        private CancellationTokenSource? _actionCts;

        #endregion

        #region 构造函数与初始化

        public MainWindow()
        {
            InitializeComponent();
            SetupUI();
        }

        /// <summary>
        /// 初始化界面数据绑定及环境
        /// </summary>
        private void SetupUI()
        {
            StepList.ItemsSource = _task.Steps;
            if (!Directory.Exists(_configDir)) Directory.CreateDirectory(_configDir);
            
            // 注册生命周期自动清理
            this.Closed += CleanUpResources;
        }

        private void CleanUpResources(object? sender, EventArgs e)
        {
            UnsetMouseHook();
            _actionCts?.Cancel();
        }

        #endregion

        #region 窗口连接逻辑

        /// <summary>
        /// 搜索并绑定模拟器窗口
        /// </summary>
        private void OnConnectClick(object sender, RoutedEventArgs e)
        {
            string windowTitle = TitleTxt.Text;
            _targetHwnd = Win32Helper.FindWindow(null, windowTitle);

            if (_targetHwnd == IntPtr.Zero)
            {
                ShowWarning("未探测到指定窗口，请检查标题。");
                return;
            }

            // 尝试查找核心渲染层句柄
            _inputHwnd = Win32Helper.FindWindowEx(_targetHwnd, IntPtr.Zero, "TheRender", null);
            if (_inputHwnd == IntPtr.Zero) 
                _inputHwnd = Win32Helper.FindWindowEx(_targetHwnd, IntPtr.Zero, "RenderWindow", null);
            
            // 保底使用主窗口
            if (_inputHwnd == IntPtr.Zero) _inputHwnd = _targetHwnd;

            ShowInfo("成功锁定输入窗口 handle: " + _inputHwnd.ToInt32().ToString("X"));
        }

        #endregion

        #region 录制引擎

        /// <summary>
        /// 录制按钮逻辑切换
        /// </summary>
        private void OnRecordToggleClick(object sender, RoutedEventArgs e)
        {
            if (_inputHwnd == IntPtr.Zero)
            {
                ShowWarning("请先成功连接模拟器后再试。");
                return;
            }

            if (!_isRecordingAttribute)
            {
                StartRecordingFlow();
            }
            else
            {
                StopRecordingFlow();
            }
        }

        private void StartRecordingFlow()
        {
            _isRecordingAttribute = true;
            RecordBtn.Content = "⏹ 停止录制";
            RecordBtn.Background = System.Windows.Media.Brushes.DimGray;
            _lastActionTimestamp = DateTime.Now;
            SetMouseHook();
        }

        private void StopRecordingFlow()
        {
            _isRecordingAttribute = false;
            RecordBtn.Content = "● 开始录制";
            RecordBtn.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 0, 0));
            UnsetMouseHook();
        }

        private void SetMouseHook()
        {
            _mouseProcDelegate = OnGlobalMouseInput;
            using (Process currentProcess = Process.GetCurrentProcess())
            using (ProcessModule currentModule = currentProcess.MainModule!)
            {
                _mouseHookHwnd = Win32Helper.SetWindowsHookEx(Win32Helper.WH_MOUSE_LL, _mouseProcDelegate, Win32Helper.GetModuleHandle(currentModule.ModuleName!), 0);
            }
        }

        private void UnsetMouseHook()
        {
            if (_mouseHookHwnd != IntPtr.Zero)
            {
                Win32Helper.UnhookWindowsHookEx(_mouseHookHwnd);
                _mouseHookHwnd = IntPtr.Zero;
            }
        }

        private IntPtr OnGlobalMouseInput(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int messageType = (int)wParam;
                if (messageType == Win32Helper.WM_LBUTTONDOWN)
                {
                    _lastDownPosition = Marshal.PtrToStructure<Win32Helper.POINT>(lParam);
                    _lastDownTimestamp = DateTime.Now;
                    _isLeftButtonActive = true;
                }
                else if (messageType == Win32Helper.WM_LBUTTONUP && _isLeftButtonActive)
                {
                    _isLeftButtonActive = false;
                    var upPoint = Marshal.PtrToStructure<Win32Helper.POINT>(lParam);
                    AddRecordedStep(_lastDownPosition, upPoint, _lastDownTimestamp);
                }
            }
            return Win32Helper.CallNextHookEx(_mouseHookHwnd, nCode, wParam, lParam);
        }

        private void AddRecordedStep(Win32Helper.POINT rawDown, Win32Helper.POINT rawUp, DateTime startTime)
        {
            var clientDown = rawDown;
            var clientUp = rawUp;
            
            if (Win32Helper.ScreenToClient(_inputHwnd, ref clientDown))
            {
                Win32Helper.ScreenToClient(_inputHwnd, ref clientUp);
                this.Dispatcher.Invoke(() => 
                {
                    var currentTime = DateTime.Now;
                    int gapDelay = (int)(startTime - _lastActionTimestamp).TotalMilliseconds;
                    int moveDuration = (int)(currentTime - startTime).TotalMilliseconds;
                    
                    // 偏移判断：5px 判定为点击，否则为滑动
                    bool flagDrag = Math.Abs(rawDown.X - rawUp.X) > 5 || Math.Abs(rawDown.Y - rawUp.Y) > 5;
                    
                    _task.Steps.Add(new ActionStep { 
                        Type = flagDrag ? ActionType.Drag : ActionType.Click,
                        X = clientDown.X, 
                        Y = clientDown.Y, 
                        EndX = clientUp.X,
                        EndY = clientUp.Y,
                        DurationMs = moveDuration,
                        DelayMs = Math.Max(0, gapDelay), 
                        Description = flagDrag ? "执行滑动" : "快速点击", 
                        Timestamp = startTime 
                    });
                    _lastActionTimestamp = currentTime;
                });
            }
        }

        #endregion

        #region 回放执行逻辑

        /// <summary>
        /// 执行当前保存的任务链
        /// </summary>
        private async void OnExecuteTaskClick(object sender, RoutedEventArgs e)
        {
            if (_inputHwnd == IntPtr.Zero || !_task.Steps.Any()) return;
            
            _actionCts = new CancellationTokenSource();
            UpdateExecutionUI(true);
            
            try 
            {
                bool isLoopMode = LoopChk.IsChecked ?? false;
                int loopGap = int.TryParse(LoopDelayTxt.Text, out int result) ? result : 1000;
                
                do 
                {
                    foreach (var step in _task.Steps) 
                    {
                        await Task.Delay(step.DelayMs, _actionCts.Token);
                        if (step.Type == ActionType.Click) 
                            await SimulateClick(step.X, step.Y);
                        else 
                            await SimulateDrag(step.X, step.Y, step.EndX, step.EndY, step.DurationMs);
                        
                        if (_actionCts.IsCancellationRequested) break;
                    }
                    if (isLoopMode) await Task.Delay(loopGap, _actionCts.Token);
                } while (isLoopMode && !_actionCts.IsCancellationRequested);
            } 
            catch (OperationCanceledException) { /* 正常取消任务处理 */ }
            finally 
            {
                UpdateExecutionUI(false);
            }
        }

        private async Task SimulateClick(int targetX, int targetY)
        {
            IntPtr posLParam = Win32Helper.MakeLParam(targetX, targetY);
            Win32Helper.PostMessage(_inputHwnd, (uint)Win32Helper.WM_MOUSEMOVE, IntPtr.Zero, posLParam);
            await Task.Delay(15);
            Win32Helper.PostMessage(_inputHwnd, (uint)Win32Helper.WM_LBUTTONDOWN, (IntPtr)1, posLParam);
            await Task.Delay(50);
            Win32Helper.PostMessage(_inputHwnd, (uint)Win32Helper.WM_LBUTTONUP, IntPtr.Zero, posLParam);
        }

        private async Task SimulateDrag(int startX, int startY, int endX, int endY, int totalDuration)
        {
            IntPtr startLp = Win32Helper.MakeLParam(startX, startY);
            IntPtr endLp = Win32Helper.MakeLParam(endX, endY);
            
            Win32Helper.PostMessage(_inputHwnd, (uint)Win32Helper.WM_MOUSEMOVE, IntPtr.Zero, startLp);
            await Task.Delay(15);
            Win32Helper.PostMessage(_inputHwnd, (uint)Win32Helper.WM_LBUTTONDOWN, (IntPtr)1, startLp);
            
            int segmentCount = 12; // 精确分段提升平滑度
            int delayPerSegment = totalDuration / segmentCount;
            for (int k = 1; k <= segmentCount; k++) 
            {
                int nextX = startX + (endX - startX) * k / segmentCount;
                int nextY = startY + (endY - startY) * k / segmentCount;
                Win32Helper.PostMessage(_inputHwnd, (uint)Win32Helper.WM_MOUSEMOVE, (IntPtr)1, Win32Helper.MakeLParam(nextX, nextY));
                await Task.Delay(Math.Max(5, delayPerSegment));
            }
            
            Win32Helper.PostMessage(_inputHwnd, (uint)Win32Helper.WM_LBUTTONUP, IntPtr.Zero, endLp);
        }

        private void UpdateExecutionUI(bool running)
        {
            RunBtn.IsEnabled = !running;
            StopTaskBtn.Visibility = running ? Visibility.Visible : Visibility.Collapsed;
        }

        private void OnStopExecutionClick(object sender, RoutedEventArgs e) => _actionCts?.Cancel();

        #endregion

        #region 存储与清理

        private void OnSaveConfigClick(object sender, RoutedEventArgs e)
        {
            if (!_task.Steps.Any()) return;
            var savePicker = new Microsoft.Win32.SaveFileDialog { InitialDirectory = _configDir, Filter = "自动化流程|*.json", FileName = "录制任务.json" };
            if (savePicker.ShowDialog() == true)
            {
                File.WriteAllText(savePicker.FileName, JsonConvert.SerializeObject(_task, Formatting.Indented));
                ConfigDisplayName.Text = Path.GetFileName(savePicker.FileName);
            }
        }

        private void OnLoadConfigClick(object sender, RoutedEventArgs e)
        {
            var openPicker = new Microsoft.Win32.OpenFileDialog { InitialDirectory = _configDir, Filter = "自动化流程|*.json" };
            if (openPicker.ShowDialog() == true)
            {
                var deserialized = JsonConvert.DeserializeObject<AutomationTask>(File.ReadAllText(openPicker.FileName));
                if (deserialized != null) 
                {
                    _task.Steps.Clear(); 
                    foreach (var stepItem in deserialized.Steps) _task.Steps.Add(stepItem);
                    ConfigDisplayName.Text = Path.GetFileName(openPicker.FileName);
                }
            }
        }

        private void OnClearStepsClick(object sender, RoutedEventArgs e) 
        { 
            _task.Steps.Clear(); 
            ConfigDisplayName.Text = "列表已重置"; 
        }

        #endregion

        #region 通用反馈封装

        private void ShowInfo(string msg) => System.Windows.MessageBox.Show(msg, "系统信息", MessageBoxButton.OK, MessageBoxImage.Information);
        private void ShowWarning(string msg) => System.Windows.MessageBox.Show(msg, "操作警告", MessageBoxButton.OK, MessageBoxImage.Warning);

        #endregion
    }
}