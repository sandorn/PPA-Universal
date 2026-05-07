using System;
using System.Reflection;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Business.Services
{
    /// <summary>
    /// 撤销/重做服务实现
    /// 
    /// 【工作原理】
    /// PowerPoint: 连续的 COM 操作会被自动合并为一个撤销条目。
    /// 调用 StartNewUndoEntry() 创建撤销边界，使每个操作成为独立的撤销条目。
    /// 
    /// WPS: 需要使用 BeginUndoGroup / EndUndoGroup 配对调用来合并操作。
    /// 使用 CreateUndoScope() 创建 IDisposable 作用域来自动管理配对调用。
    /// </summary>
    public class UndoService : IUndoService
    {
        private readonly ILogger _logger;

        public UndoService(ILogger logger)
        {
            _logger = logger ?? NullLogger.Instance;
        }

        public bool StartNewUndoEntry(IApplicationContext context, string undoEntryName)
        {
            // 此方法仅用于 PowerPoint
            // WPS 应使用 CreateUndoScope() 方法
            if (context == null || context.Platform != PlatformType.PowerPoint)
            {
                return false;
            }

            try
            {
                var nativeApp = context.NativeApplication;
                if (nativeApp == null) return false;

                return TryStartNewUndoEntryPowerPoint(nativeApp, undoEntryName);
            }
            catch (Exception ex)
            {
                _logger.LogDebug($"设置撤销边界失败: {ex.Message}");
                return false;
            }
        }

        public void EndUndoEntry(IApplicationContext context)
        {
            // PowerPoint 不需要显式结束
            // WPS 应使用 CreateUndoScope() 的 Dispose 来结束
        }

        public IDisposable CreateUndoScope(IApplicationContext context, string undoEntryName)
        {
            if (context == null)
            {
                return new NullUndoScope();
            }

            switch (context.Platform)
            {
                case PlatformType.PowerPoint:
                    // PowerPoint: 调用 StartNewUndoEntry 创建边界
                    TryStartNewUndoEntryPowerPoint(context.NativeApplication, undoEntryName);
                    return new NullUndoScope(); // PowerPoint 不需要结束调用
                
                case PlatformType.WPS:
                    // WPS: 目前无法实现撤销合并（BeginUndoGroup 不支持，StartNewUndoEntry 无效）
                    // 详见文件末尾 TODO
                    return new NullUndoScope();
                
                default:
                    return new NullUndoScope();
            }
        }

        public bool IsSupported(IApplicationContext context)
        {
            if (context == null) return false;
            return context.Platform == PlatformType.PowerPoint || 
                   context.Platform == PlatformType.WPS;
        }

        private bool TryStartNewUndoEntryPowerPoint(object nativeApp, string undoEntryName)
        {
            if (nativeApp == null) return false;

            // 尝试通过 dynamic 调用
            try
            {
                dynamic dynamicApp = nativeApp;
                dynamicApp.StartNewUndoEntry();
                _logger.LogDebug($"已设置撤销边界 (PowerPoint): {undoEntryName}");
                return true;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // dynamic 调用失败，尝试反射
            }

            // 尝试通过反射调用
            var appType = nativeApp.GetType();
            var method = appType.GetMethod("StartNewUndoEntry",
                BindingFlags.Public | BindingFlags.Instance,
                null,
                Type.EmptyTypes,
                null);

            if (method != null)
            {
                method.Invoke(nativeApp, null);
                _logger.LogDebug($"已设置撤销边界 (PowerPoint, 反射): {undoEntryName}");
                return true;
            }

            _logger.LogDebug("PowerPoint 未找到 StartNewUndoEntry 方法");
            return false;
        }
    }

    /// <summary>
    /// 空撤销作用域 - 用于不支持撤销合并的平台
    /// </summary>
    internal class NullUndoScope : IDisposable
    {
        public void Dispose() { }
    }

    /*
     * TODO: WPS 撤销合并技术路线总结
     * 
     * 1. StartNewUndoEntry: 
     *    - 调用不报错，但实际上没有任何效果，撤销记录仍然是零散的。
     *    
     * 2. BeginUndoGroup / EndUndoGroup:
     *    - 运行时抛出 Exception: “System.__ComObject”未包含“BeginUndoGroup”的定义。
     *    - 说明 WPS 的 COM 接口并未通过自动化暴露此方法。
     *    
     * 3. ScreenUpdating = false:
     *    - 禁用屏幕更新未能实现隐式的撤销合并。
     *    
     * 结论：目前 WPS 平台暂时无法实现完美的撤销合并（原子操作）。
     * 后续如果 WPS 开放了相关 API 或有新的黑科技（如通过 CommandBars 执行宏），可在此处扩展 WpsUndoScope 实现。
     */
}

