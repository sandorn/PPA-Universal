using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;
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
            UniversalIntegration.Logger?.LogInformation("Ribbon loaded successfully");
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

        #region 表格格式化

        public void OnFormatThreeLineTable(object control)
        {
            try
            {
                var shapes = GetSelectedShapes();
                if (shapes == null || !shapes.Any())
                {
                    ShowMessage("请先选择包含表格的形状");
                    return;
                }

                var tableShapes = shapes.Where(s => s?.IsTable == true && s.Table != null).ToList();
                if (tableShapes.Count == 0)
                {
                    ShowMessage("选中形状中没有表格");
                    return;
                }

                var formatService = UniversalIntegration.GetService<ITableFormatService>();
                if (formatService == null)
                {
                    ShowMessage("表格格式化服务不可用");
                    return;
                }

                foreach (var shape in tableShapes)
                {
                    formatService.FormatTableAsThreeLine(shape.Table);
                }

                UniversalIntegration.Logger?.LogInformation($"已对 {tableShapes.Count} 个表格应用三线表格式");
            }
            catch (Exception ex)
            {
                UniversalIntegration.Logger?.LogError($"三线表格式化失败: {ex.Message}", ex);
                ShowMessage($"三线表格式化失败: {ex.Message}");
            }
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
            UniversalIntegration.Logger?.LogInformation($"Alignment reference changed to: {_currentReference}");
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
                UniversalIntegration.Logger?.LogInformation($"Alignment executed: {alignmentType}, Reference: {_currentReference}");
            }
            catch (Exception ex)
            {
                UniversalIntegration.Logger?.LogError($"Alignment failed: {ex.Message}", ex);
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
                UniversalIntegration.Logger?.LogInformation($"Distribution executed: {distributionType}");
            }
            catch (Exception ex)
            {
                UniversalIntegration.Logger?.LogError($"Distribution failed: {ex.Message}", ex);
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
                UniversalIntegration.Logger?.LogInformation("Size operation executed");
            }
            catch (Exception ex)
            {
                UniversalIntegration.Logger?.LogError($"Size operation failed: {ex.Message}", ex);
                ShowMessage($"尺寸操作失败: {ex.Message}");
            }
        }

        private List<IShapeContext> GetSelectedShapes()
        {
            try
            {
                var context = UniversalIntegration.Context;
                var selection = context?.Selection;
                
                if (selection == null)
                {
                    return null;
                }

                // 检查是否有形状选择
                if (selection.Type != SelectionType.Shapes && selection.ShapeCount == 0)
                {
                    return null;
                }

                return selection.SelectedShapes?.ToList();
            }
            catch (Exception ex)
            {
                UniversalIntegration.Logger?.LogError($"GetSelectedShapes failed: {ex.Message}", ex);
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

        #endregion
    }
}
