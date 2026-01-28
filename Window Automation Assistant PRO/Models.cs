using System;
using System.Collections.ObjectModel;

namespace WindowTool
{
    /// <summary>
    /// 定义自动化动作的分类
    /// </summary>
    public enum ActionType
    {
        /// <summary>普通点击动作</summary>
        Click,
        /// <summary>鼠标长按并平滑移动动作</summary>
        Drag
    }

    /// <summary>
    /// 表示自动化序列中的单一步骤
    /// </summary>
    public class ActionStep
    {
        /// <summary>执行的操作类型</summary>
        public ActionType Type { get; set; } = ActionType.Click;

        /// <summary>起始 X 坐标</summary>
        public int X { get; set; }
        
        /// <summary>起始 Y 坐标</summary>
        public int Y { get; set; }

        /// <summary>结束 X 坐标（仅 Drag 类型有效）</summary>
        public int EndX { get; set; }

        /// <summary>结束 Y 坐标（仅 Drag 类型有效）</summary>
        public int EndY { get; set; }

        /// <summary>动作执行时长（毫秒，用于模拟平滑滑动）</summary>
        public int DurationMs { get; set; } = 50;

        /// <summary>相较于上一个动作的等待延时（毫秒）</summary>
        public int DelayMs { get; set; } = 500;

        /// <summary>步骤描述</summary>
        public string Description { get; set; } = "未描述动作";

        /// <summary>记录时的时间戳</summary>
        public DateTime Timestamp { get; set; }
    }

    /// <summary>
    /// 包含完整自动化流程的任务模型
    /// </summary>
    public class AutomationTask
    {
        /// <summary>任务名称</summary>
        public string Name { get; set; } = "自动执行任务";

        /// <summary>动作步骤集合</summary>
        public ObservableCollection<ActionStep> Steps { get; set; } = new();
    }
}