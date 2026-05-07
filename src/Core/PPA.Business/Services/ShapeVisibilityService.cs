using System.Collections.Generic;
using System.Linq;
using PPA.Business.Abstractions;
using PPA.Core.Abstraction;
using PPA.Logging;

namespace PPA.Business.Services
{
    /// <summary>
    /// 形状可见性服务实现
    /// </summary>
    public class ShapeVisibilityService : IShapeVisibilityService
    {
        private readonly ILogger _logger;
        private readonly IShapeOperations _shapeOps;

        public ShapeVisibilityService(ILogger logger, IShapeOperations shapeOps)
        {
            _logger = logger ?? NullLogger.Instance;
            _shapeOps = shapeOps;
        }

        public void HideShapes(IEnumerable<IShapeContext> shapes)
        {
            var shapeList = shapes?.ToList();
            if (shapeList == null || shapeList.Count == 0)
            {
                _logger.LogWarning("没有选中任何形状");
                return;
            }

            _logger.LogInformation($"隐藏 {shapeList.Count} 个形状");

            foreach (var shape in shapeList)
            {
                if (shape?.NativeShape != null)
                {
                    _shapeOps.SetVisible(shape.NativeShape, false);
                }
            }

            _logger.LogInformation("隐藏形状完成");
        }

        public void ShowAllHiddenShapes(ISlideContext slide)
        {
            if (slide == null)
            {
                _logger.LogWarning("幻灯片为空");
                return;
            }

            _logger.LogInformation("显示所有隐藏的形状");

            int hiddenCount = 0;
            foreach (var shape in slide.Shapes)
            {
                if (shape?.NativeShape != null)
                {
                    bool isVisible = _shapeOps.GetVisible(shape.NativeShape);
                    if (!isVisible)
                    {
                        _shapeOps.SetVisible(shape.NativeShape, true);
                        hiddenCount++;
                    }
                }
            }

            _logger.LogInformation($"已显示 {hiddenCount} 个隐藏的形状");
        }
    }
}

