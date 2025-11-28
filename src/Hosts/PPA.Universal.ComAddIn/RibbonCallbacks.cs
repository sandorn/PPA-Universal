using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Universal.Integration;

namespace PPA.Universal.ComAddIn
{
    /// <summary>
    /// Ribbon 回调处理类
    /// </summary>
    [ComVisible(true)]
    public class RibbonCallbacks
    {
        private object _ribbon;
        private AlignmentReference _currentReference = AlignmentReference.SelectedObjects;

        /// <summary>
        /// Ribbon 加载时调用
        /// </summary>
        public void Ribbon_OnLoad(object ribbon)
        {
            _ribbon = ribbon;
            Log("Ribbon loaded successfully");
        }

        #region 对齐操作

        public void OnAlignLeft(object control)
        {
            ExecuteAlignment(AlignmentType.Left);
        }

        public void OnAlignRight(object control)
        {
            ExecuteAlignment(AlignmentType.Right);
        }

        public void OnAlignTop(object control)
        {
            ExecuteAlignment(AlignmentType.Top);
        }

        public void OnAlignBottom(object control)
        {
            ExecuteAlignment(AlignmentType.Bottom);
        }

        public void OnAlignCenterH(object control)
        {
            ExecuteAlignment(AlignmentType.CenterHorizontal);
        }

        public void OnAlignCenterV(object control)
        {
            ExecuteAlignment(AlignmentType.CenterVertical);
        }

        #endregion

        #region 分布操作

        public void OnDistributeH(object control)
        {
            ExecuteDistribution(DistributionType.Horizontal);
        }

        public void OnDistributeV(object control)
        {
            ExecuteDistribution(DistributionType.Vertical);
        }

        #endregion

        #region 尺寸操作

        public void OnEqualWidth(object control)
        {
            ExecuteSizeOperation(s => s.SetEqualWidth(GetSelectedShapes()));
        }

        public void OnEqualHeight(object control)
        {
            ExecuteSizeOperation(s => s.SetEqualHeight(GetSelectedShapes()));
        }

        public void OnEqualSize(object control)
        {
            ExecuteSizeOperation(s => s.SetEqualSize(GetSelectedShapes()));
        }

        #endregion

        #region 参考选项

        public void OnAlignRefChanged(object control, string selectedId, int selectedIndex)
        {
            _currentReference = selectedIndex switch
            {
                0 => AlignmentReference.SelectedObjects,
                1 => AlignmentReference.Slide,
                2 => AlignmentReference.FirstObject,
                3 => AlignmentReference.LastObject,
                _ => AlignmentReference.SelectedObjects
            };
            Log($"Alignment reference changed to: {_currentReference}");
        }

        public int GetAlignRefIndex(object control)
        {
            return _currentReference switch
            {
                AlignmentReference.SelectedObjects => 0,
                AlignmentReference.Slide => 1,
                AlignmentReference.FirstObject => 2,
                AlignmentReference.LastObject => 3,
                _ => 0
            };
        }

        #endregion

        #region 辅助方法

        private void ExecuteAlignment(AlignmentType alignmentType)
        {
            try
            {
                var shapes = GetSelectedShapes();
                if (shapes == null || !shapes.Any())
                {
                    ShowMessage("请先选择要对齐的形状");
                    return;
                }

                var service = UniversalIntegration.GetService<IAlignmentService>();
                if (service == null)
                {
                    ShowMessage("对齐服务不可用");
                    return;
                }

                service.Align(shapes, alignmentType, _currentReference);
                Log($"Alignment executed: {alignmentType}, Reference: {_currentReference}");
            }
            catch (Exception ex)
            {
                Log($"Alignment failed: {ex.Message}");
                ShowMessage($"对齐操作失败: {ex.Message}");
            }
        }

        private void ExecuteDistribution(DistributionType distributionType)
        {
            try
            {
                var shapes = GetSelectedShapes();
                if (shapes == null || shapes.Count < 3)
                {
                    ShowMessage("分布操作需要选择至少 3 个形状");
                    return;
                }

                var service = UniversalIntegration.GetService<IAlignmentService>();
                if (service == null)
                {
                    ShowMessage("对齐服务不可用");
                    return;
                }

                service.Distribute(shapes, distributionType);
                Log($"Distribution executed: {distributionType}");
            }
            catch (Exception ex)
            {
                Log($"Distribution failed: {ex.Message}");
                ShowMessage($"分布操作失败: {ex.Message}");
            }
        }

        private void ExecuteSizeOperation(Action<IAlignmentService> operation)
        {
            try
            {
                var shapes = GetSelectedShapes();
                if (shapes == null || shapes.Count < 2)
                {
                    ShowMessage("尺寸操作需要选择至少 2 个形状");
                    return;
                }

                var service = UniversalIntegration.GetService<IAlignmentService>();
                if (service == null)
                {
                    ShowMessage("对齐服务不可用");
                    return;
                }

                operation(service);
                Log("Size operation executed");
            }
            catch (Exception ex)
            {
                Log($"Size operation failed: {ex.Message}");
                ShowMessage($"尺寸操作失败: {ex.Message}");
            }
        }

        private List<IShapeContext> GetSelectedShapes()
        {
            try
            {
                var context = UniversalIntegration.Context;
                var selection = context?.Selection;

                if (selection == null || selection.Type != SelectionType.Shapes)
                {
                    return null;
                }

                return selection.SelectedShapes?.ToList();
            }
            catch (Exception ex)
            {
                Log($"GetSelectedShapes failed: {ex.Message}");
                return null;
            }
        }

        private void ShowMessage(string message)
        {
            System.Windows.Forms.MessageBox.Show(
                message, 
                "PPA Universal", 
                System.Windows.Forms.MessageBoxButtons.OK, 
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private void Log(string message)
        {
            try
            {
                var logPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PPA.Universal",
                    "Ribbon.log");

                var directory = System.IO.Path.GetDirectoryName(logPath);
                if (!System.IO.Directory.Exists(directory))
                {
                    System.IO.Directory.CreateDirectory(directory);
                }

                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
                System.IO.File.AppendAllText(logPath, line);
            }
            catch
            {
                // 忽略日志写入失败
            }
        }

        #endregion
    }
}
