using PPA.Core.Abstraction;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Adapter.PowerPoint
{
    /// <summary>
    /// PowerPoint 窗口上下文实现
    /// </summary>
    public class PowerPointWindowContext : IWindowContext
    {
        private readonly NETOP.DocumentWindow _window;

        public PowerPointWindowContext(NETOP.DocumentWindow window)
        {
            _window = window;
        }

        public ViewType ViewType
        {
            get
            {
                try
                {
                    var viewType = _window?.ViewType;
                    if (viewType == null) return ViewType.Normal;

                    // 使用数值比较以避免 NetOffice 枚举版本差异
                    var viewTypeInt = (int)viewType.Value;
                    switch (viewTypeInt)
                    {
                        case 1: // ppViewNormal
                        case 9: // ppViewNormal variants
                            return ViewType.Normal;
                        case 6: // ppViewOutline
                            return ViewType.Outline;
                        case 4: // ppViewSlideSorter
                            return ViewType.SlideSorter;
                        case 5: // ppViewNotesPage
                            return ViewType.NotesPage;
                        case 7: // ppViewTitleMaster
                        case 2: // ppViewSlideMaster
                        case 3: // ppViewHandoutMaster
                        case 8: // ppViewNotesMaster
                            return ViewType.Master;
                        default:
                            return ViewType.Normal;
                    }
                }
                catch
                {
                    return ViewType.Normal;
                }
            }
        }

        public ISlideContext ActiveSlide
        {
            get
            {
                try
                {
                    var view = _window?.View;
                    var slide = view?.Slide as NETOP.Slide;
                    return slide != null ? new PowerPointSlideContext(slide) : null;
                }
                catch
                {
                    return null;
                }
            }
        }

        public int Zoom
        {
            get
            {
                try { return _window?.View?.Zoom ?? 100; }
                catch { return 100; }
            }
            set
            {
                try { if (_window?.View != null) _window.View.Zoom = value; }
                catch { }
            }
        }

        public object NativeWindow => _window;
    }
}
